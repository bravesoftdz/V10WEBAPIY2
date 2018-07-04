unit VerDPaq;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, Hctrls, HSysMenu, ExtCtrls, Ent1, HEnt1, HStatus,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  hmsgbox ;

procedure DatePaquet;

type
  TFVerDPaq = class(TForm)
    Panel1: TPanel;
    HMTrad: THSystemMenu;
    FLettrage: THValComboBox;
    HPB: TPanel;
    BValider: THBitBtn;
    BFerme: THBitBtn;
    FEche: TCheckBox;
    FModePaie: THValComboBox;
    TFModePaie: THLabel;
    FChangeMp: TCheckBox;
    MsgRien: THMsgBox;
    procedure BValiderClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FLettrageChange(Sender: TObject);
    procedure FChangeMpClick(Sender: TObject);
  private
    { Déclarations privées }
    procedure PrepareShow;
  public
    { Déclarations publiques }

  Procedure RepartNLettre ;
  Procedure LigneNumeche ;
  end;

implementation

{$R *.DFM}

uses CPTESAV, UtilPgi;

procedure DatePaquet;
var VDPaq : TFVerDPaq ;
BEGIN
if Not _BlocageMonoPoste(True,'',TRUE) then Exit ;
VDPaq:=TFVerDPaq.Create(Application) ;
try
  // GCO - 24/04/2002
  if VH^.EnSerie then
  begin
    VDPaq.PrepareShow;
    VDPaq.BValider.Click;
  end
  else
    VDPaq.ShowModal ;
  // FIN GCO

 finally
  VDPaq.Free ;
  _DeblocageMonoPoste(True,'',TRUE) ;
 end ;
Screen.Cursor:=SyncrDefault ;
END ;

// GCO - 24/04/2002
procedure TFVerDPaq.PrepareShow;
begin
  FLettrage.ItemIndex:=0 ;
  FModePaie.ItemIndex:=0 ;
  FLettrageChange(Nil) ;
end;

procedure TFVerDPaq.FormShow(Sender: TObject);
begin
  PrepareShow;
end;
// Fin GCO

procedure TFVerDPaq.BValiderClick(Sender: TObject);
begin
// GCO - 24/04/2002
if not VH^.EnSerie then
begin
  if MsgRien.Execute(0,'','')<>mryes then exit ;
end;
// Fin GCO
EnableControls(Self,False) ;
If FLettrage.ItemIndex<>1 then RepartNLettre ;
Case FLettrage.ItemIndex of
  0 : SAVDATEMINDATEMAXPAQUET(Tous) ;
  1 : SAVDATEMINDATEMAXPAQUET(Lettre) ;
  2 : SAVDATEMINDATEMAXPAQUET(NonLettre) ;
  End ;
if FEche.Checked then LigneNumeche ;
EnableControls(Self,True) ;
end;

procedure TFVerDPaq.BFermeClick(Sender: TObject);
begin Close ; end;

Procedure TFVerDPaq.RepartNLettre ;
Var OkPourLet, OkPourPoint, Lettrable, PointBqe, LigneOk : Boolean ;
    Q : TQuery ; St : String ; NbMvt : Integer ;
BEGIN
begintrans ;
InitMove(100,'') ; NbMvt:=0 ;
Q:=OpenSQL(' Select E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, '+
           ' E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE, E_GENERAL, E_AUXILIAIRE, '+
           ' E_ETATLETTRAGE, G_POINTABLE, T_LETTRABLE, G_NATUREGENE '+
           ' From ECRITURE '+
           ' LEFT JOIN GENERAUX On G_GENERAL=E_GENERAL '+
           ' LEFT JOIN TIERS On T_AUXILIAIRE=E_AUXILIAIRE '+
           ' where e_lettrage="" '+
           ' order BY E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE ',True) ;
While Not Q.EOF do
   BEGIN
   MoveCur(FALSE) ;
   OkPourLet:=(Q.Fields[9].AsString='AL')
            or(Q.Fields[9].AsString='PL')
            or (Q.Fields[9].AsString='TL') ;
   OkPourPoint:=(Q.Fields[9].AsString='RI') ;
   Lettrable:=(Q.Fields[11].AsString='X') ;
   PointBqe:=(Q.Fields[10].AsString='X')and(Q.Fields[12].AsString='BQE') ;
   St:='' ; LigneOk:=True ;
   if Lettrable then if Not OkPourLet then begin St:='UPDATE ECRITURE set E_EtatLettrage="AL", ' ; LigneOk:=False ; end ;
   if PointBqe then  if Not OkPourPoint then begin St:='UPDATE ECRITURE set E_EtatLettrage="RI", ' ; LigneOk:=False ; end ;
   if not LigneOk then if St<>'' then
      BEGIN
     If FChangeMp.checked then St:=St+' E_MODEPAIE="'+FModePaie.Value+'", ' ;
     St:=St+'e_couverture=0, e_couverturedev=0 '+
             'where e_auxiliaire="'+Q.FindField('E_AUXILIAIRE').AsString+'" and e_general="'+Q.FindField('E_GENERAL').AsString+'" '+
             'and E_JOURNAL="'+Q.FindField('E_JOURNAL').AsString+'" and E_EXERCICE="'+Q.FindField('E_EXERCICE').AsString+'" '+
             'and e_datecomptable="'+USDateTime(Q.FindField('e_datecomptable').AsDateTime)+'" and E_NUMEROPIECE='+IntToStr(Q.FindField('E_NUMEROPIECE').AsInteger)+' '+
             'and e_numligne='+IntToStr(Q.FindField('e_numLigne').AsInteger)+' and E_NUMECHE='+IntToStr(Q.FindField('E_NUMECHE').AsInteger)+
             'and E_QUALIFPIECE="'+Q.FindField('E_QUALIFPIECE').AsString+'" ' ;
      ExecuteSql(St) ;
      END ;
   Q.Next ;
   if NbMvt>=100 then BEGIN NbMvt:=0 ; FiniMove ; InitMove(100,'') ; END ;
   END ;
Ferme(Q) ;
committrans ; FiniMove ;
END ;

Procedure TFVerDPaq.LigneNumeche ;
Var OkPourLet, OkPourPoint, Lettrable, PointBqe, LigneOk : Boolean ;
    Q : TQuery ; St : String ; NbMvt : Integer ;
BEGIN
begintrans ;
InitMove(100,'') ; NbMvt:=0 ;
Q:=OpenSQL(' Select E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, '+
           ' E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE, E_GENERAL, E_AUXILIAIRE, '+
           ' E_ETATLETTRAGE, G_POINTABLE, T_LETTRABLE, G_NATUREGENE '+
           ' From ECRITURE '+
           ' LEFT JOIN GENERAUX On G_GENERAL=E_GENERAL '+
           ' LEFT JOIN TIERS On T_AUXILIAIRE=E_AUXILIAIRE '+
           ' where e_eche="X" '+
           ' order BY E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE ',True) ;
While Not Q.Eof do
      BEGIN
      MoveCur(FALSE) ;
      OkPourLet:=(Q.Fields[9].AsString='AL')or(Q.Fields[9].AsString='PL')
               or(Q.Fields[9].AsString='TL') ;
      OkPourPoint:=(Q.Fields[9].AsString='RI') ;
      Lettrable:=(Q.Fields[11].AsString='X') ;
      PointBqe:=(Q.Fields[10].AsString='X')and(Q.Fields[12].AsString='BQE') ;
      St:='' ; LigneOk:=True ;
      if Lettrable then if OkPourLet then LigneOk:=(Q.Fields[5].AsInteger>=1) ;
      if PointBqe then  if OkPourPoint then LigneOk:=(Q.Fields[5].AsInteger>=1) ;
      if not LigneOk then if St='' then
         BEGIN
         St:=' Update ECRITURE Set e_numeche=1 '+
                'where e_auxiliaire="'+Q.FindField('E_AUXILIAIRE').AsString+'" and e_general="'+Q.FindField('E_GENERAL').AsString+'" '+
                'and E_JOURNAL="'+Q.FindField('E_JOURNAL').AsString+'" and E_EXERCICE="'+Q.FindField('E_EXERCICE').AsString+'" '+
                'and e_datecomptable="'+USDateTime(Q.FindField('e_datecomptable').AsDateTime)+'" and E_NUMEROPIECE='+IntToStr(Q.FindField('E_NUMEROPIECE').AsInteger)+' '+
                'and e_numligne='+IntToStr(Q.FindField('e_numLigne').AsInteger)+' and E_NUMECHE='+IntToStr(Q.FindField('E_NUMECHE').AsInteger)+
                'and E_QUALIFPIECE="'+Q.FindField('E_QUALIFPIECE').AsString+'"' ;
         ExecuteSql(St) ;
         END ;
      Q.Next ;
      if NbMvt>=100 then BEGIN NbMvt:=0 ; FiniMove ; InitMove(100,'') ; END ;
      END ;
Ferme(Q) ;
committrans ; FiniMove ;
END ;

procedure TFVerDPaq.FLettrageChange(Sender: TObject);
begin
FChangeMp.Enabled:=(FLettrage.ItemIndex<>1) ;
if FLettrage.ItemIndex<>1 then FChangeMpClick(Nil) ;
end;

procedure TFVerDPaq.FChangeMpClick(Sender: TObject);
begin
TFModePaie.Enabled:=FChangeMp.Checked ;
FModePaie.Enabled:=FChangeMp.Checked ;
end;

end.
