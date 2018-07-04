{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 27/01/2005
Modifié le ... :   /  /    
Description .. : Passage en eAGL
Mots clefs ... : 
*****************************************************************}
unit MajCodBu;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, Hctrls, hmsgbox, HSysMenu, StdCtrls, Buttons, ExtCtrls, Ent1, HEnt1,
{$IFDEF EAGLCLIENT}
  UTob,
{$ELSE}
  DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF} 
{$ENDIF}
  HStatus;

Procedure MajCode(Li : TList ; Table : String ) ;

Type TCodBud = Class
      UnCod : String ;
      UnLib : String ;
      UnAbr : String ;
      UnSig : String ;
      UnSen : String ;
      UnRub : String ;
      UnCpR : String ;
      UnAxe : String ;
      UnCat : String ;
      End ;

type
  TFMajCodBu = class(TForm)
    HPB: TPanel;
    Panel1: TPanel;
    BAide: THBitBtn;
    BFerme: THBitBtn;
    BValider: THBitBtn;
    HMTrad: THSystemMenu;
    HM: THMsgBox;
    FListe: THGrid;
    BAnnuler: THBitBtn;
    procedure BFermeClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FListeKeyPress(Sender: TObject; var Key: Char);
    procedure FListeCellEnter(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    procedure BValiderClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  private
    Li : TList ;
    Table : String ;
    Cle,CompteRub,Pref : String ;
    Modifier : Boolean ;
    WMinX,WMinY    : Integer ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure RempliFListe ;
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

Procedure MajCode(Li : TList ; Table : String ) ;
var FMajCodBu : TFMajCodBu ;
BEGIN
if (Li=Nil) or (Li.Count=0) or (Table='') then Exit ;
FMajCodBu:=TFMajCodBu.Create(Application) ;
  Try
   FMajCodBu.Li:=Li ;
   FMajCodBu.Table:=Table ;
   FMajCodBu.ShowModal ;
  Finally
   FMajCodBu.Free ;
  End ;
SourisNormale ;
END ;

procedure TFMajCodBu.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFMajCodBu.WMGetMinMaxInfo(var MSG: Tmessage);
BEGIN with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end; END ;

procedure TFMajCodBu.FormCreate(Sender: TObject);
begin WMinX:=Width ; WMinY:=Height ; Modifier:=False ; end;

procedure TFMajCodBu.FormShow(Sender: TObject);
begin
if Table='BUDSECT' then
   BEGIN
   Caption:=HM.Mess[3] ; FListe.Cells[2,0]:=HM.Mess[1] ;
   Cle:='BS_BUDSECT' ; CompteRub:='BS_SECTIONRUB' ; Pref:='BS_' ;
   END else
   BEGIN
   Caption:=HM.Mess[2] ; FListe.Cells[2,0]:=HM.Mess[0] ;
   Cle:='BG_BUDGENE' ; CompteRub:='BG_COMPTERUB' ; Pref:='BG_' ;
   END ;
RempliFListe ;
UpdateCaption(Self) ;
end;

Procedure TFMajCodBu.RempliFListe ;
Var i : Integer ;
    X : TCodBud ;
BEGIN
FListe.VidePile(False) ; Li.Pack ;
for i:=0 to Li.Count-1 do
  BEGIN
  if Li.Items[i]=Nil then Continue ;
  X:=Li.Items[i] ;
  FListe.Cells[0,FListe.RowCount-1]:=X.UnCod ;
  FListe.Cells[1,FListe.RowCount-1]:=X.UnLib ;
  FListe.Cells[2,FListe.RowCount-1]:=X.UnCpR ;
  FListe.Cells[3,FListe.RowCount-1]:=X.UnRub ;
  FListe.RowCount:=FListe.RowCount+1 ;
  END ;
if FListe.RowCount>2 then FListe.RowCount:=FListe.RowCount-1 ;
END ;

procedure TFMajCodBu.FListeKeyPress(Sender: TObject; var Key: Char);
begin
if Li.Count=0 then BEGIN Key:=#0 ; Exit ; END ;
if (FListe.Col=0) or (FListe.Col=3) then Key:=UpCase(Key) ;
end;

procedure TFMajCodBu.FListeCellEnter(Sender: TObject; var ACol,ARow: Longint; var Cancel: Boolean);
begin if FListe.Col=2 then if ACol>2 then FListe.Col:=1 else FListe.Col:=3 ; end;

procedure TFMajCodBu.BValiderClick(Sender: TObject);
Var QLoc : TQuery ;
    CodVide,DoublonCod : Boolean ;
    Sql,St : String ;
    i : Integer ;
begin
if FListe.Cells[0,1]='' then Exit ;
if Li.Count=0 then Exit ;
if HM.Execute(4,'','')<>mrYes then Exit ;
if Table='BUDGENE' then Sql:='Select * from BUDGENE Where BG_BUDGENE="'+W_W+'"'
                   else Sql:='Select * from BUDSECT Where BS_BUDSECT="'+W_W+'"' ;
DoublonCod:=False ; CodVide:=False ;
QLoc:=OpenSql(Sql,False) ;
InitMove(FListe.RowCount-1,'') ;
for i:=1 to Fliste.RowCount-1 do
   BEGIN
   MoveCur(False) ;
   if (FListe.Cells[0,i]='') or (FListe.Cells[1,i]='') or (FListe.Cells[3,i]='') then BEGIN CodVide:=True ; Continue ; END ;
   if Presence(Table,Cle,Copy(FListe.Cells[0,i],1,17)) or
      Presence(Table,Pref+'RUB',Copy(FListe.Cells[3,i],1,5))then BEGIN DoublonCod:=True ; Continue ; END ;
   QLoc.Insert ; InitNew(QLoc) ;
   QLoc.FindField(Cle).AsString:=Copy(FListe.Cells[0,i],1,17) ;
   QLoc.FindField(Pref+'LIBELLE').AsString:=Copy(FListe.Cells[1,i],1,35) ;
   QLoc.FindField(Pref+'ABREGE').AsString:=Copy(FListe.Cells[1,i],1,17) ;
   QLoc.FindField(Pref+'SIGNE').AsString:=TCodBud(Li.Items[i-1]).UnSig ;
   QLoc.FindField(Pref+'SENS').AsString:=TCodBud(Li.Items[i-1]).UnSen ;
   QLoc.FindField(Pref+'RUB').AsString:=FListe.Cells[3,i] ;
   St:=TCodBud(Li.Items[i-1]).UnCpR ; if Pos(';;',St)<>Length(St)-1 then St:=St+';;' ;
   QLoc.FindField(CompteRub).AsString:=St ;
   if Table='BUDSECT' then
      BEGIN
      QLoc.FindField(Pref+'AXE').AsString:=TCodBud(Li.Items[i-1]).UnAxe ;
      QLoc.FindField(Pref+'CATEGORIE').AsString:=TCodBud(Li.Items[i-1]).UnCat ;
      END ;
   QLoc.Post ;
   TCodBud(Li.Items[i-1]).Free ; Li.Items[i-1]:=Nil ;
   END ;
Ferme(QLoc) ;
if DoublonCod then HM.Execute(5,'','') ;
if CodVide    then HM.Execute(6,'','') ;
FiniMove ; RempliFListe ; Modifier:=False ;
end;

procedure TFMajCodBu.FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin if Li.Count=0 then Exit ; Modifier:=True ; end;

procedure TFMajCodBu.BAnnulerClick(Sender: TObject);
begin
if Not Modifier then Exit ;
if HM.Execute(7,'','')<>mrYes then Exit ;
RempliFListe ; Modifier:=False ;
end;

procedure TFMajCodBu.BAideClick(Sender: TObject);
begin CallHelpTopic(Self) ; end;

end.
