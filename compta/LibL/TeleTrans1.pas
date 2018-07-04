unit Teletrans ;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, Grids, Hctrls, ComCtrls, EtbUser,
  DBCtrls, Mask, DB, DBTables, Ent1, HEnt1, Hqry, hmsgbox, PrintDBG,
  BanqueCP,HStatus, HDB, HSysMenu, HTB97, ed_tools, HPanel, UiUtil ;

Type T_TypeTele = (ttlEmission,ttlReception,ttlIntegration) ; 

Procedure LanceTeletrans ( LeMode : T_TypeTele ) ;

type
  TFTeletrans = class(TForm)
    DockTop: TDock97;
    DockRight: TDock97;
    DockLeft: TDock97;
    HDiv: THMsgBox;
    HMTrad: THSystemMenu;
    G: THGrid;
    Dock971: TDock97;
    DockBottom: TDock97;
    Valide97: TToolbar97;
    BValide: TToolbarButton97;
    BAnnuler: TToolbarButton97;
    BAide: TToolbarButton97;
    BImprimer: TToolbarButton97;
    BChercher: TToolbarButton97;
    PEntete: THPanel;
    HLabel1: THLabel;
    cDest: THValComboBox;
    RRecept: TRadioGroup;
    REmet: TRadioGroup;
    cCopie: TCheckBox;
    BAjout: TToolbarButton97;
    AppelF: TOpenDialog;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BValideClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure cDestChange(Sender: TObject);
    procedure REmetClick(Sender: TObject);
    procedure RReceptClick(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure GMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure GDblClick(Sender: TObject);
    procedure BAjoutClick(Sender: TObject);
  private
    WMinX,WMinY,GX,GY : Integer ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
    Function  RempliLigne ( Q : TQuery ) : boolean ; 
    procedure CocheDecoche ( ARow : integer ; Next : Boolean ) ;
    procedure ConcatFichiers ( sSource,sDest : String ; Premier : boolean ) ;
    procedure Emissions ;
    procedure SetLesVisibles ;
    procedure Receptions ;
  public
    NeedRech : boolean ;
    LeMode   : T_TypeTele ;
  end;

implementation

{$R *.DFM}

Procedure LanceTeletrans ( LeMode : T_TypeTele ) ;
Var PP : THPanel ;
    X  : TFTeletrans ;
BEGIN
PP:=FindInsidePanel ;
X:=TFTeletrans.Create(Application) ;
X.LeMode:=LeMode ;
if PP=Nil then
   BEGIN
    Try
     X.ShowModal ;
    Finally
     X.Free ;
    End ;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;
END ;

procedure TFTeletrans.FormClose(Sender: TObject; var Action: TCloseAction);
begin
G.VidePile(True) ;
if isInside(Self) then Action:=caFree ;
end;

procedure TFTeletrans.SetLesVisibles ;
BEGIN
Case LeMode of
     ttlEmission : BEGIN
                   REmet.Visible:=True ; RRecept.Visible:=False ;
                   cCopie.Visible:=True ;
                   Caption:=HDiv.Mess[4] ;
                   BValide.Hint:=HDiv.Mess[7] ; BValide.Caption:=HDiv.Mess[7] ;
                   END ;
    ttlReception : BEGIN
                   REmet.Visible:=False ; RRecept.Visible:=True ;
                   cCopie.Visible:=False ; BChercher.Visible:=False ;
                   Caption:=HDiv.Mess[5] ;
                   BValide.Hint:=HDiv.Mess[8] ; BValide.Caption:=HDiv.Mess[8] ;
                   BAjout.Visible:=False ; 
                   END ;
  ttlIntegration : BEGIN
                   REmet.Visible:=False ; RRecept.Visible:=True ;
                   cCopie.Visible:=False ;
                   Caption:=HDiv.Mess[6] ;
                   BValide.Hint:=HDiv.Mess[9] ; BValide.Caption:=HDiv.Mess[9] ;
                   END ;
   END ;
UpdateCaption(Self) ;    
END ;

procedure TFTeletrans.FormShow(Sender: TObject);
begin
z_GetDestinataires(cDest.Values,cDest.Items) ;
SetLesVisibles ;
NeedRech:=True ; G.VidePile(True) ;
G.ColWidths[3]:=0 ;
G.GetCellCanvas:=GetCellCanvas ;
end;

Procedure TFTeletrans.GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
BEGIN
if ((ARow<>0) And (G.Cells[G.ColCount-1,ARow]='+'))
   then G.Canvas.Font.Style:=G.Canvas.Font.Style+[fsItalic]
   else G.Canvas.Font.Style:=G.Canvas.Font.Style-[fsItalic] ;
END ;

procedure TFTeletrans.BImprimerClick(Sender: TObject);
begin PrintDBGrid(G,Nil,Caption,'') ; End;

procedure TFTeletrans.ConcatFichiers ( sSource,sDest : String ; Premier : boolean ) ;
Var FSource,FDest : TextFile ;
    io : integer ;
    St : String ;
BEGIN
AssignFile(FDest,sDest) ;
{$I-} if Premier then Rewrite(FDest) else Append(FDest) ; {$I+}
if IoResult<>0 then BEGIN {$I-} CloseFile(FDest) ; {$I+} io:=ioResult ; Exit ; END ;
AssignFile(FSource,sSource) ;
{$I-} Reset(FSource) ; {$I+}
if IoResult<>0 then BEGIN {$I-} CloseFile(FSource) ; {$I+} io:=ioResult ; Exit ; END ;
While Not EOF(FSource) do
   BEGIN
   Readln(FSource,St) ;
   WriteLn(FDest,St) ;
   END ;
CloseFile(FDest) ; CloseFile(FSource) ;
END ;

procedure TFTeletrans.Receptions ;
Var ie : integer ; 
BEGIN
//ie:=z_Teletransmission(PChar(NomFichier),PChar(Chemin),PChar(Carte),PDest,0) ;
END ;

procedure TFTeletrans.Emissions ;
Var Lig,ie : integer ;
    Premier,Okok : boolean ;
    Carte,NomFichier,Chemin,StF : String ;
    PDest : PChar ;
BEGIN
Premier:=True ; Carte:='' ; NomFichier:='' ; Okok:=False ;
for Lig:=1 to G.RowCount-1 do if G.Cells[G.ColCount-1,Lig]='+' then
    BEGIN
    if Premier then
       BEGIN
       Chemin:=ExtractFilePath(G.Cells[0,Lig]) ;
       if Chemin<>'' then if Chemin[Length(Chemin)]<>'\' then Chemin:=Chemin+'\' ;
       NomFichier:=V_PGI.USER+'_TEMP' ;
       END ;
    ConcatFichiers(G.Cells[0,Lig],Chemin+NomFichier,Premier) ;
    Premier:=False ; Okok:=True ;
    END ;
if Not Okok then BEGIN HDiv.Execute(2,'','') ; Exit ; END ;
if NomFichier='' then Exit ;
Case REmet.ItemIndex of
   0 : Carte:='E000' ;
   1 : Carte:='E001' ;
   2 : Carte:='E002' ;
   else Exit ;
   END ;
if cDest.Value='' then PDest:=Nil else PDest:=PChar(cDest.Value) ;
ie:=z_Teletransmission(PChar(NomFichier),PChar(Chemin),PChar(Carte),PDest,0) ;
if ie<>0 then
   BEGIN
   HDiv.Execute(3,'','') ;
   END else
   BEGIN
   for Lig:=1 to G.RowCount-1 do if G.Cells[G.ColCount-1,Lig]='+' then
       BEGIN
       StF:=G.Cells[0,Lig] ;
       if cCopie.Checked then
          BEGIN
          DeleteFile(StF+'_') ; RenameFile(StF,StF+'_') ;
          END else
          BEGIN
          DeleteFile(StF) ;
          END ;
       END ;
   END ;
if NomFichier<>'' then DeleteFile(Chemin+NomFichier) ;
END ;

procedure TFTeletrans.BValideClick(Sender: TObject);
begin
if ((NeedRech) and (LeMode<>ttlreception)) then BEGIN HDiv.Execute(0,'','') ; Exit ; END ;
Case LeMode of
      ttlEmission : BEGIN
                    if HDiv.Execute(1,'','')<>mrYes then Exit ;
                    Emissions ; BChercherClick(Nil) ;
                    END ;
     ttlReception : BEGIN
                    if HDiv.Execute(10,'','')<>mrYes then Exit ;
                    Receptions ;
                    END ;
   ttlIntegration : BEGIN
                    if HDiv.Execute(11,'','')<>mrYes then Exit ;
                    END ;
   END ;
end;

procedure TFTeletrans.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFTeletrans.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFTeletrans.BAnnulerClick(Sender: TObject);
begin
if Not isInside(Self) then Close else CloseInsidePanel(Self) ;
end;

Function TFTeletrans.RempliLigne ( Q : TQuery ) : boolean ;
Var St,Codebanque,Libelle : String ;
BEGIN
Result:=False ;
if LeMode=ttlEmission then
   BEGIN
   Case REmet.ItemIndex of
      0 : St:=Q.FindField('BQ_REPLCR').AsString ;
      1 : St:=Q.FindField('BQ_REPVIR').AsString ;
      2 : St:=Q.FindField('BQ_REPPRELEV').AsString ;
      END ;
   END else
   BEGIN
   St:=Q.FindField('BQ_REPRELEVE').AsString ;
   END ;
CodeBanque:=Q.FindField('BQ_GENERAL').AsString ;
Libelle:=Q.FindField('BQ_LIBELLE').AsString ;
if FileExists(St) then
   BEGIN
   Result:=True ; 
   G.Cells[0,G.RowCount-1]:=St ;
   G.Cells[1,G.RowCount-1]:=CodeBanque ;
   G.Cells[2,G.RowCount-1]:=Libelle ;
   G.RowCount:=G.RowCount+1 ;
   END ;
END ;

procedure TFTeletrans.BChercherClick(Sender: TObject);
Var Q : TQuery ;
    Okok : boolean ;
begin
G.VidePile(True) ; Okok:=False ;
Q:=OpenSQL('SELECT * from BANQUECP WHERE BQ_DESTINATAIRE="'+cDest.Value+'"',True) ;
While Not Q.EOF do
   BEGIN
   if RempliLigne(Q) then Okok:=True ;
   Q.Next ;
   END ;
Ferme(Q) ;
if Okok then G.RowCount:=G.RowCount-1 ;
NeedRech:=False ;
end;

procedure TFTeletrans.cDestChange(Sender: TObject);
begin
NeedRech:=True ;
end;

procedure TFTeletrans.REmetClick(Sender: TObject);
begin
NeedRech:=True ;
end;

procedure TFTeletrans.RReceptClick(Sender: TObject);
begin
NeedRech:=True ;
end;

procedure TFTeletrans.FormResize(Sender: TObject);
begin
G.ColWidths[3]:=0 ;
end;

procedure TFTeletrans.CocheDecoche ( ARow : integer ; Next : Boolean ) ;
BEGIN
if G.Cells[G.ColCount-1,ARow]='+' then G.Cells[G.ColCount-1,ARow]:=' ' else G.Cells[G.ColCount-1,ARow]:='+' ;
if ((ARow=G.RowCount-1) and (Next)) then Next:=False ;
if Next then G.Row:=ARow+1 ;
G.Invalidate ;
END ;

procedure TFTeletrans.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var OkG,Vide : Boolean ;
begin
OkG:=G.Focused ; Vide:=(Shift=[]) ;
Case Key of
   VK_SPACE  : if ((OkG) and (Vide)) then CocheDecoche(G.Row,False) ;
   END ;
end;

procedure TFTeletrans.GMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
Var C,R : Longint ;
begin
GX:=X ; GY:=Y ;
if ((ssCtrl in Shift) and (Button=mbLeft)) then
   BEGIN
   G.MouseToCell(X,Y,C,R) ;
   if R>0 then CocheDecoche(G.Row,False) ;
   END ;

end;

procedure TFTeletrans.GDblClick(Sender: TObject);
Var C,R : longint ;
begin
G.MouseToCell(GX,GY,C,R) ;
if R>0 then CocheDecoche(G.Row,True) ;
end;

procedure TFTeletrans.BAjoutClick(Sender: TObject);
Var i : integer ;
begin
if AppelF.Execute then
   BEGIN
   for i:=1 to G.RowCount-1 do if G.Cells[0,i]=AppelF.FileName then Exit ;
   if G.Cells[0,G.RowCount-1]<>'' then G.RowCount:=G.RowCount+1 ;
   G.Cells[0,G.RowCount-1]:=AppelF.FileName ;
   G.Cells[G.ColCount-1,G.RowCount-1]:='+' ;
   END ;
end;

end.
