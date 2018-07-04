unit AuditS3;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, hmsgbox, HSysMenu, Ent1, HEnt1, {$IFNDEF DBXPRESS}dbtables,
  HTB97{$ELSE}uDbxDataSet{$ENDIF},
  Hctrls, HTB97, HPanel, UiUtil, ParamSoc ;

Function LanceAuditPourClotureS3 : Boolean ;

type
  TFAuditS3 = class(TForm)
    Panel1: TPanel;
    HPB: TToolWindow97;
    BAide: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    V1: TLabel;
    B1: TToolbarButton97;
    V2: TLabel;
    B2: TToolbarButton97;
    V3: TLabel;
    B3: TToolbarButton97;
    V4: TLabel;
    B4: TToolbarButton97;
    V5: TLabel;
    B5: TToolbarButton97;
    V6: TLabel;
    B6: TToolbarButton97;
    Panel2: TPanel;
    R1: TLabel;
    R2: TLabel;
    R3: TLabel;
    R4: TLabel;
    R5: TLabel;
    R6: TLabel;
    MsgRien: THMsgBox;
    MsgLibel: THMsgBox;
    HMTrad: THSystemMenu;
    Dock: TDock97;
    procedure FormShow(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure B1Click(Sender: TObject);
    procedure B2Click(Sender: TObject);
    procedure B3Click(Sender: TObject);
    procedure B4Click(Sender: TObject);
    procedure B5Click(Sender: TObject);
    procedure B6Click(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Déclarations privées }
    StR1, StR2, StR3,
    StR4, StR5, StR6 : String ;
    OnSort : Boolean ;
    Fait1, Fait2, Fait3,
    Fait4, Fait5, Fait6 : Boolean ;
    Procedure MajSociete ;
    Procedure ControleEdtLegal ;
    Procedure ControlePerValid ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

uses VerCpta,
{$IFNDEF CCS3}
     VerCai,QRLegal,
{$ENDIF}
     VerSolde, CloPerio, Cloture,  Valperio ;

Function LanceAuditPourClotureS3 : Boolean ;
var Aud : TFAuditS3 ;
    OutProg : Boolean ;
    PP : THPanel ;
BEGIN
Result:=FALSE ;
Aud:=TFAuditS3.Create(Application) ; OutProg:=FALSE ;
Aud.OnSort:=OutProg ;
PP:=FindInsidePanel ;
if ((PP=Nil) or (True)) then
   BEGIN
    try
     Aud.ShowModal ;
    finally
     OutProg:=Aud.OnSort ;
     Aud.Free ;
    end ;
   If OutProg Then Result:=TRUE ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(Aud,PP) ;
   Aud.Show ;
   END ;
END ;

procedure TFAuditS3.FormShow(Sender: TObject);
Var Q : TQuery ; StEtat, StRecup : String ; i : Integer ;
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
StEtat:='' ; StRecup:='' ;
{$IFDEF SPEC302}
Q:=OpenSql(' Select SO_RECUPCPTA from SOCIETE where SO_SOCIETE="'+V_PGI.CodeSociete+'" ',True) ;
if Not Q.Eof then StRecup:=Q.fields[0].AsString ;
Ferme(Q) ;
{$ELSE}
StRecup:=GetParamSoc('SO_RECUPCPTA') ;
{$ENDIF}
For i:=1 to Length(StRecup) do
    BEGIN
    if (StRecup[i]='*')or(StRecup[i]='-')or(StRecup[i]='+') then StEtat:=StEtat+StRecup[i] else StEtat:=StEtat+'*' ;
    END ;
if (Length(StEtat)>=6) then
   BEGIN
   StR1:=StEtat[1] ; StR2:=StEtat[2] ; StR3:=StEtat[3] ;
   StR4:=StEtat[4] ; StR5:=StEtat[5] ; StR6:=StEtat[6] ;
   If StR1='-' then R1.Caption:=MsgLibel.Mess[2] else if StR1='+' then R1.Caption:=MsgLibel.Mess[1] else R1.Caption:=MsgLibel.Mess[0] ;
   If StR2='-' then R2.Caption:=MsgLibel.Mess[2] else if StR2='+' then R2.Caption:=MsgLibel.Mess[1] else R2.Caption:=MsgLibel.Mess[0] ;
   If StR3='-' then R3.Caption:=MsgLibel.Mess[2] else if StR3='+' then R3.Caption:=MsgLibel.Mess[1] else R3.Caption:=MsgLibel.Mess[0] ;
   If StR4='-' then R4.Caption:=MsgLibel.Mess[2] else if StR4='+' then R4.Caption:=MsgLibel.Mess[1] else R4.Caption:=MsgLibel.Mess[0] ;
   If StR5='-' then R5.Caption:=MsgLibel.Mess[2] else if StR5='+' then R5.Caption:=MsgLibel.Mess[1] else R5.Caption:=MsgLibel.Mess[0] ;
   If StR6='-' then R6.Caption:=MsgLibel.Mess[2] else if StR6='+' then R6.Caption:=MsgLibel.Mess[1] else R6.Caption:=MsgLibel.Mess[0] ;
   END Else
   BEGIN
   StR1:='*' ; R1.Caption:=MsgLibel.Mess[0]  ;
   StR2:='*' ; R2.Caption:=MsgLibel.Mess[0]  ;
   StR3:='*' ; R3.Caption:=MsgLibel.Mess[0]  ;
   StR4:='*' ; R4.Caption:=MsgLibel.Mess[0]  ;
   StR5:='*' ; R5.Caption:=MsgLibel.Mess[0]  ;
   StR6:='*' ; R6.Caption:=MsgLibel.Mess[0]  ;
   END ;
Fait1:=False ; Fait2:=False ; Fait3:=False ;
Fait4:=False ; Fait5:=False ; Fait6:=False ;
if EstSerie(S5) then V6.Caption:=MSGLibel.Mess[3] ;
{$IFDEF CCS3}
B5.Visible:=False ; V5.Visible:=False ; 
{$ENDIF}
end;

procedure TFAuditS3.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFAuditS3.BValiderClick(Sender: TObject);
begin
MajSociete ;
// GP le 31/01/2001 if ClotureComptable(True) then OnSort:=TRUE ;
if ctxPCL in V_PGI.PGIContexte then onSort := True
else
begin
  if ClotureComptable(True) then OnSort:=TRUE ;
end;
end;

procedure TFAuditS3.B1Click(Sender: TObject);
begin
if VerPourClo(VH^.EnCours) then
   BEGIN
   R1.Caption:=MsgLibel.Mess[1] ;
   StR1:='+' ;
   END Else
   BEGIN
   R1.Caption:=MsgLibel.Mess[2] ;
   StR1:='-' ;
   END ;
end;

procedure TFAuditS3.B2Click(Sender: TObject);
begin
ControlePerValid ;
end;

procedure TFAuditS3.B3Click(Sender: TObject);
begin
ControleEdtLegal ;
end;

procedure TFAuditS3.B4Click(Sender: TObject);
begin
ControleSolde ;
R4.Caption:=MsgLibel.Mess[1] ; StR4:='+' ;
end;

procedure TFAuditS3.B5Click(Sender: TObject);
begin
{$IFNDEF CCS3}
Case CtrlCaiClo of
  cPasFait : BEGIN if Not Fait5 then begin R5.Caption:=MsgLibel.Mess[0] ; StR5:='*' ; end ; END ;
  cOk      : BEGIN R5.Caption:=MsgLibel.Mess[1] ; StR5:='+' ; Fait5:=True ; END ;
  cPasOk   : BEGIN R5.Caption:=MsgLibel.Mess[2] ; StR5:='-' ; Fait5:=True ; END ;
  End ;
{$ENDIF}
end;

procedure TFAuditS3.B6Click(Sender: TObject);
begin
Case CtrlTo(VH^.EnCours) of
   cOk      : BEGIN R6.Caption:=MsgLibel.Mess[1] ; StR6:='+' ; Fait6:=True ; END ;
   cPasFait : BEGIN if Not Fait6 then begin R6.Caption:=MsgLibel.Mess[0] ; StR6:='*' ; end ; END ;
  End ;
end;

Procedure TFAuditS3.MajSociete ;
Var St, StSql : String ;
BEGIN
St:=StR1+StR2+StR3+StR4+StR5+StR6 ;
{$IFDEF SPEC302}
StSql:=' UPDATE SOCIETE Set SO_RECUPCPTA="'+St+'" where SO_SOCIETE="'+V_PGI.CodeSociete+'" ' ;
ExecuteSql(StSql) ;
{$ELSE}
SetParamSoc('SO_RECUPCPTA',St) ;
{$ENDIF}
END ;

Procedure TFAuditS3.ControleEdtLegal ;
Var Q : TQuery ;
BEGIN
{$IFDEF CCS3}
R3.Caption:=MsgLibel.Mess[1] ; StR3:='+' ;
{$ELSE}
Q:=OpenSql(' select * from edtlegal where ed_obligatoire="X" and ed_exercice="'+VH^.Encours.Code+'" '+
           ' and (ED_TYPEEDITION<>"GLT" and ED_TYPEEDITION<>"GLG" and ED_TYPEEDITION<>"BLT" '+
           ' and ED_TYPEEDITION<>"BLG" and ED_TYPEEDITION<>"JLD") ',True) ;
If Q.Eof then
   BEGIN
   if (MsgRien.Execute(1,caption,'')<>mryes) then BEGIN Ferme(Q) ; Exit ; END ;
   END Else
   BEGIN
   Fait3:=True ;
   R3.Caption:=MsgLibel.Mess[1] ; StR3:='+' ;
   Ferme(Q) ; MsgRien.Execute(3,caption,'') ; Exit ;
   END ;
Ferme(Q) ;
Case LanceEdtLegClo of
  cPasFait : BEGIN if Not Fait3 then begin R3.Caption:=MsgLibel.Mess[0] ; StR3:='*' ; end ; END ;
  cOk      : BEGIN R3.Caption:=MsgLibel.Mess[1] ; StR3:='+' ; END ;
  cPasOk   : BEGIN R3.Caption:=MsgLibel.Mess[2] ; StR3:='-' ; END ;
  End ;
{$ENDIF}
END ;

Procedure TFAuditS3.ControlePerValid ;
Var Q : TQuery ; i, NbPer : Byte ; StValide : String ;
    Valide : Boolean ;
BEGIN
StValide:='' ; NbPer:=VH^.Encours.NombrePeriode ; Valide:=False ;
Q:=OpenSql('Select EX_VALIDEE from exercice where ex_exercice="'+VH^.Encours.Code+'" ',True) ;
if Not Q.Eof then StValide:=Q.fields[0].AsString ;
Ferme(Q) ;
If StValide='' then Exit ;
For i:=1 to NbPer do
    BEGIN
    If StValide[i]='-' then Valide:=False Else Valide:=True ; if Not Valide then Break ;
    END ;
if Valide then
   BEGIN
   R2.Caption:=MsgLibel.Mess[1] ; StR2:='+' ; MsgRien.Execute(2,caption,'') ;
   END Else
   BEGIN
   if (MsgRien.Execute(0,caption,'')<>mrYes) then Exit ;
   Case ValPerPourClo(True, False) of
     cPasFait : BEGIN if Not Fait2 then begin R2.Caption:=MsgLibel.Mess[0] ; StR2:='*' ; end ; END ;
     cOk      : BEGIN R2.Caption:=MsgLibel.Mess[1] ; StR2:='+' ; Fait2:=True ; END ;
     cPasOk   : BEGIN R2.Caption:=MsgLibel.Mess[2] ; StR2:='-' ; Fait2:=True ; END ;
     End ;
   END ;
END ;

procedure TFAuditS3.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFAuditS3.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if Parent is THPanel then Action:=caFree ;
end;

end.
