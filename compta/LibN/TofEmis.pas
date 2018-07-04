unit TofEmis ;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, Grids, Hctrls, ComCtrls, EtbUser,
  DBCtrls, Mask, DB, Ent1, HEnt1, Hqry, hmsgbox, PrintDBG,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  BANQUECP_TOM,
  HStatus, HDB, HSysMenu, HTB97, ed_tools, HPanel, UiUtil, Utof,Vierge;

type
  TOF_Teletrans = class(TOF)
    G: THGrid;
    cDest : THValComboBox ;
    REmet : TRadioGroup ;
    cCopie : TCheckBox ;
    procedure BImprimerClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure GDblClick(Sender: TObject);
    procedure BAjoutClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure GMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
  private
    GX,GY : Integer ;
    Procedure GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
    Function  RempliLigne ( Q : TQuery ) : boolean ;
    procedure CocheDecoche ( ARow : integer ; Next : Boolean ) ;
    procedure ConcatFichiers ( sSource,sDest : String ; Premier : boolean ) ;
    procedure Emissions ;
  public
    procedure Onload ; override ;
    procedure OnUpdate; override ;
  end;

const SR_NOMFICHIER =   0 ;
      SR_CODEBQE    =   1 ;
      SR_INTITULE   =   2 ;
      SR_COCHE      =   3 ;

      	// libellés des messages
	TexteMessage: array[1..10] of string 	= (
          {01}        '0;Télétransmission ETEBAC; Vous devez lancer une recherche au préalable !;W;O;O;O;'
          {02}        ,'1;Télétransmission ETEBAC; Confirmez-vous l''émission ETEBAC ?;Q;YN;Y;Y;'
          {03}        ,'2;Télétransmission ETEBAC; Vous devez sélectionner au moins un fichier !;W;O;O;O;'
          {04}        ,'3;Télétransmission ETEBAC; La télétransmission a échoué !;W;O;O;O;'
          {05}        ,'Emission par télétransmission ETEBAC des fichiers CFONB'
          {06}        ,'Emision'
          {07}        ,'10;Télétransmission ETEBAC; Confirmez-vous la réception ETEBAC ?;Q;YN;Y;Y;'
          {08}        ,'11;Télétransmission ETEBAC; Confirmez-vous l''intégration ?;Q;YN;Y;Y;'
          {09}        ,'11;Télétransmission ETEBAC; Le module Liaison Bancaire n''est pas sérialisé ;W;O;O;O;'
          {10}        ,''
                     );

implementation

procedure TOF_Teletrans.OnLoad ;
var Bimprimer,Baide,BChercher,BAjout : TButton;
begin
TFVierge(Ecran).HelpContext:=7594000 ;
REmet:=TRadioGroup(GetControl('REMET')) ;
cCopie:=TCheckBox(GetControl('CCOPIE')) ;
BImprimer:=TButton(GetControl('BIMPRIME')) ;
BAide:=TButton(GetControl('BAIDE')) ;
BAjout:=TButton(GetControl('BAJOUT')) ;
BChercher:=TButton(GetControl('BCHERCHER')) ;
G:=THGrid(GetControl('GFILES'));
cDest:=THValCombobox(GetControl('CDEST')) ;
if (BImprimer <> nil ) and (not Assigned(BImprimer.OnClick)) then BImprimer.OnClick:=BImprimerClick;
if (BAide <> nil ) and (not Assigned(BAide.OnClick)) then BAide.OnClick:=BAideClick;
if (BAjout <> nil ) and (not Assigned(BAjout.OnClick)) then BAjout.OnClick:=BAjoutClick;
if (BChercher <> nil ) and (not Assigned(BChercher.OnClick)) then BChercher.OnClick:=BChercherClick;
if G<>nil then
  begin
  G.ColWidths[SR_COCHE]:=0 ;
  G.GetCellCanvas:=GetCellCanvas ;
  if not Assigned(G.OnDblClick) then G.OnDblClick:=GDblClick;
  end;
if cDest<>nil then z_GetDestinataires(cDest.Values,cDest.Items) ;
cDest.ItemIndex:=0 ;
end;

Procedure TOF_Teletrans.GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
BEGIN
if ((ARow<>0) And (G.Cells[G.ColCount-1,ARow]='+'))
   then G.Canvas.Font.Style:=G.Canvas.Font.Style+[fsItalic]
   else G.Canvas.Font.Style:=G.Canvas.Font.Style-[fsItalic] ;
END ;

procedure TOF_Teletrans.BImprimerClick(Sender: TObject);
begin PrintDBGrid(G,Nil,'Export CFONB : Télétransmission ETEBAC','') ; End;

procedure TOF_Teletrans.ConcatFichiers ( sSource,sDest : String ; Premier : boolean ) ;
Var FSource,FDest : TextFile ;
    //io : integer ;
    St : String ;
BEGIN
AssignFile(FDest,sDest) ;
{$I-} if Premier then Rewrite(FDest) else Append(FDest) ; {$I+}
if IoResult<>0 then BEGIN {$I-} CloseFile(FDest) ; {$I+ io:=ioResult  ;} Exit ; END ;
AssignFile(FSource,sSource) ;
{$I-} Reset(FSource) ; {$I+}
if IoResult<>0 then BEGIN {$I-} CloseFile(FSource) ; {$I+ io:=ioResult  ;} Exit ; END ;
While Not EOF(FSource) do
   BEGIN
   Readln(FSource,St) ;
   WriteLn(FDest,St) ;
   END ;
CloseFile(FDest) ; CloseFile(FSource) ;
END ;

procedure TOF_Teletrans.Emissions ;
Var Lig,ie : integer ;
    Premier,Okok : boolean ;
    Carte,NomFichier,Chemin,StF : String ;
    PDest : PChar ;
BEGIN
if not VH^.OkModEtebac then begin HShowMessage(TexteMessage[9],'','') ; Exit ; end ;
Premier:=True ; Carte:='' ; NomFichier:='' ; Okok:=False ;
for Lig:=1 to G.RowCount-1 do if G.Cells[G.ColCount-1,Lig]='+' then
  begin
  if Premier then
    begin
    Chemin:=ExtractFilePath(G.Cells[0,Lig]) ;
    if Chemin<>'' then if Chemin[Length(Chemin)]<>'\' then Chemin:=Chemin+'\' ;
    NomFichier:=V_PGI.USER+'_TEMP' ;
    end ;
  ConcatFichiers(G.Cells[SR_NOMFICHIER,Lig],Chemin+NomFichier,Premier) ;
  Premier:=False ; Okok:=True ;
  end ;
if Not Okok then BEGIN HShowMessage(TexteMessage[3],'','') ; Exit ; END ;
if NomFichier='' then Exit ;
case REmet.ItemIndex of
  0 : Carte:='E000' ;
  1 : Carte:='E001' ;
  2 : Carte:='E002' ;
  else Exit ;
  end;
if cDest.Value='' then PDest:=Nil else PDest:=PChar(cDest.Value) ;
ie:=z_Teletransmission(PChar(NomFichier),PChar(Chemin),PChar(Carte),PDest,0) ;
if ie<>0 then HShowMessage(TexteMessage[4],'','')
else
  begin
  for Lig:=1 to G.RowCount-1 do
    if G.Cells[G.ColCount-1,Lig]='+' then
      begin
      StF:=G.Cells[0,Lig] ;
      if cCopie.Checked then begin DeleteFile(StF+'_') ; RenameFile(StF,StF+'_') ; end
                        else DeleteFile(StF) ;
      end ;
  end ;
if NomFichier<>'' then DeleteFile(Chemin+NomFichier) ;
end ;

procedure TOF_Teletrans.OnUpdate;
begin
if HShowMessage(TexteMessage[2],'','')<>mrYes then Exit ;
Emissions ; BChercherClick(Nil) ;
end;

procedure TOF_Teletrans.BAideClick(Sender: TObject);
begin
CallHelpTopic(Ecran) ;
end;

Function TOF_Teletrans.RempliLigne ( Q : TQuery ) : boolean ;
Var St,Codebanque,Libelle : String ;
BEGIN
Result:=False ;
case REmet.ItemIndex of
  0 : St:=Q.FindField('BQ_REPLCR').AsString ;
  1 : St:=Q.FindField('BQ_REPVIR').AsString ;
  2 : St:=Q.FindField('BQ_REPPRELEV').AsString ;
  end ;
CodeBanque:=Q.FindField('BQ_GENERAL').AsString ;
Libelle:=Q.FindField('BQ_LIBELLE').AsString ;
if FileExists(St) then
  begin
  Result:=True ;
  G.Cells[SR_NOMFICHIER,G.RowCount-1]:=St ;
  G.Cells[SR_CODEBQE,G.RowCount-1]:=CodeBanque ;
  G.Cells[SR_INTITULE,G.RowCount-1]:=Libelle ;
  G.Cells[SR_COCHE,G.RowCount-1]:='+';
  G.RowCount:=G.RowCount+1 ;
  end;
end ;

procedure TOF_Teletrans.BChercherClick(Sender: TObject);
Var Q : TQuery ;
    Okok : boolean ;
begin
G.VidePile(True) ; Okok:=False ;
Q:=OpenSQL('SELECT * FROM BANQUECP WHERE BQ_DESTINATAIRE="'+cDest.Value+'" AND BQ_NODOSSIER="'+V_PGI.NoDossier+'"',True) ; // 24/10/2006 YMO Multisociétés

While Not Q.EOF do
   BEGIN
   if RempliLigne(Q) then Okok:=True ;
   Q.Next ;
   END ;
Ferme(Q) ;
if Okok then G.RowCount:=G.RowCount-1 ;
end;

procedure TOF_Teletrans.CocheDecoche ( ARow : integer ; Next : Boolean ) ;
BEGIN
if G.Cells[SR_COCHE,ARow]='+' then G.Cells[SR_COCHE,ARow]:=' ' else G.Cells[SR_COCHE,ARow]:='+' ;
if ((ARow=G.RowCount-1) and (Next)) then Next:=False ;
if Next then G.Row:=ARow+1 ;
G.Invalidate ;
END ;

procedure TOF_Teletrans.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var OkG,Vide : Boolean ;
begin
OkG:=G.Focused ; Vide:=(Shift=[]) ;
Case Key of
   VK_SPACE  : if ((OkG) and (Vide)) then CocheDecoche(G.Row,False) ;
   END ;
end;

procedure TOF_Teletrans.GMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
Var C,R : Longint ;
begin
GX:=X ; GY:=Y ;
if ((ssCtrl in Shift) and (Button=mbLeft)) then
   BEGIN
   G.MouseToCell(X,Y,C,R) ;
   if R>0 then CocheDecoche(G.Row,False) ;
   END ;

end;

procedure TOF_Teletrans.GDblClick(Sender: TObject);
//Var C,R : longint ;
begin
{G.MouseToCell(GX,GY,C,R) ;
if R>0 then}
CocheDecoche(G.Row,True) ;
end;


procedure TOF_Teletrans.BAjoutClick(Sender: TObject);
var i,j : integer ;OpenDialog : TOpenDialog ;Okok : Boolean;
Begin
if not VH^.OkModEtebac then begin HShowMessage(TexteMessage[9],'','') ; Exit ; end ;
OpenDialog:=TOpenDialog.create(Ecran);
OpenDialog.InitialDir:='C:\';
OpenDialog.Filter:=TraduireMemoire('Tous les fichiers (*.*)')+'|*.*';
if OpenDialog.execute then
  begin
  for j:=0 to OpenDialog.Files.Count-1 do
    begin
    Okok:=TRUE;
    for i:=1 to G.RowCount-1 do if G.Cells[SR_NOMFICHIER,i]=OpenDialog.Files.Strings[j] then Okok:=FALSE ;
    if Okok then
      begin
      if G.Cells[SR_NOMFICHIER,G.RowCount-1]<>'' then G.RowCount:=G.RowCount+1 ;
      G.Cells[SR_NOMFICHIER,G.RowCount-1]:=OpenDialog.FileName ;
      G.Cells[SR_COCHE,G.RowCount-1]:='+' ;
      end ;
    end;
  end;
OpenDialog.free;
end;

Initialization
registerclasses([TOF_Teletrans]);
end.
