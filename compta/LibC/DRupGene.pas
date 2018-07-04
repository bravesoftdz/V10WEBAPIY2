unit DRupGene;

interface

uses              
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, Hctrls, StdCtrls, Buttons, ExtCtrls,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  hmsgbox, Hent1, HStatus,
  Ent1,
{$IFDEF EAGLCLIENT}
  uTob,
  UtileAGL, // PrintDBGrid
{$ELSE}
  PrintDBG, // PrintDBGrid
{$ENDIF}
  HSysMenu ;

Type
TabBool= Array[1..6]of Boolean ;
TabLib = Array[1..6]of String ;

Procedure DetailPlanRuptureGene(RupClass : TabBool ; LibClass : TabLib ; Nature,Plan : String) ;

type
  TFDrupgene = class(TForm)
    Pbouton  : TPanel;
    FAutoSave: TCheckBox;
    PTop     : TPanel;
    TPlan    : THLabel;
    TLibplan : THLabel;
    Eplan    : TEdit;
    ElibPlan : TEdit;
    Fliste   : THGrid;
    Edit1    : TEdit;
    MsgBox   : THMsgBox;
    Panel1: TPanel;
    BTag: THBitBtn;
    BAide: THBitBtn;
    BFerme: THBitBtn;
    BValider: THBitBtn;
    BImprimer: THBitBtn;
    Bdetag: THBitBtn;
    HMTrad: THSystemMenu;
    Nb1: TLabel;
    Tex1: TLabel;
    procedure BFermeClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BdetagClick(Sender: TObject);
    procedure BTagClick(Sender: TObject);
    procedure FlisteMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure BValiderClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure FlisteKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  private
    RupClass : TabBool ;
    LibClass : TabLib ;
    Plan     : String ;
    Nature   : String ;
    ListCod  : TStringList ;
    WMinX,WMinY : Integer ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure Libelle ;
    Procedure TagDetag(Avec : Boolean) ;
    Procedure RempliListe ;
    Procedure RempliGrid ;
    Procedure FaitCaption ;
    Procedure BrancheHelpContext ;
    Procedure CompteElemSelectionner ;
    Function  ListeVide : Boolean ;
    Procedure GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
    Procedure InverseSelection ;
  public
    { Déclarations publiques }
  end;

Type
   UneChaine = Class
       StLib : String ;
   END ;

implementation

{$R *.DFM}

uses UtilPgi ;

procedure TFDrupgene.BFermeClick(Sender: TObject);
begin Close ; end;

Procedure TFDrupgene.FaitCaption ;
BEGIN
if Nature='RUG' then Caption:=Caption+' '+MsgBox.Mess[3]
                else Caption:=Caption+' '+MsgBox.Mess[4] ;
UpdateCaption(Self) ;
END ;

procedure TFDrupgene.FormShow(Sender: TObject);
Var i : LongInt ;
begin
FListe.GetCellCanvas:=GetCellCanvas ;
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
FListe.ColWidths[FListe.ColCount-1]:=0 ;
Libelle ; FaitCaption ; BrancheHelpContext ;
ListCod:=TStringList.Create ; ListCod.Sorted:=True ; ListCod.Duplicates:=DupIgnore ;
RempliListe ; RempliGrid ;
for i:=0 to ListCod.Count-1 do
    ListCod.Objects[i].Free ;
ListCod.Free ;
if FListe.RowCount>2 then FListe.RowCount:=FListe.RowCount-1 ;
CompteElemSelectionner ;
FListe.SetFocus ;
end;

Procedure TFDrupgene.RempliListe ;
Var Q : TQuery ;
    i : Byte ;
    St : String ;
    Pst : UneChaine ;
    Sql : String ;
BEGIN
if Nature='RUG' then Sql:='Select G_GENERAL,G_LIBELLE From GENERAUX Order by G_GENERAL'
                else Sql:='Select T_AUXILIAIRE,T_LIBELLE From TIERS Order by T_AUXILIAIRE' ;
Q:=OpenSql(Sql,True) ;
While Not Q.Eof do
   BEGIN
   for i:=1 to 6 do
     BEGIN
     if RupClass[i]=True then
        BEGIN
        if LibClass[i]<>''then St:=(Q.Fields[1].AsString+' '+LibClass[i])
                          else St:=(Q.Fields[1].AsString) ;
        Pst:=UneChaine.Create ; Pst.StLib:=St ;
        ListCod.AddObject(Copy(Q.Fields[0].AsString,1,i)+'x',Pst) ;
        END ;
     END ;
   Q.Next ;
   END ;
Ferme(Q) ;
END ;

Procedure TFDrupgene.RempliGrid ;
Var i : LongInt ;
BEGIN
InitMove(ListCod.Count,MsgBox.Mess[2]) ;
for i:=0 to ListCod.Count-1 do
   BEGIN
   Fliste.Cells[0,FListe.RowCount-1]:=ListCod.Strings[i] ;
   Fliste.Cells[1,FListe.RowCount-1]:=UneChaine(ListCod.Objects[i]).StLib ;
   FListe.Cells[FListe.ColCount-1,FListe.RowCount-1]:='*' ;
   FListe.RowCount:=FListe.RowCount+1 ; MoveCur(False) ;
   END ;
FiniMove ;
END ;

Procedure TFDrupgene.Libelle ;
Var Q : TQuery ;
    i : Byte ;
    St : String ;
BEGIN
Q:=OpenSql('Select CC_LIBELLE From CHOIXCOD Where CC_TYPE="'+Nature+'" And CC_CODE="'+Plan+'"',True) ;
ElibPlan.Text:=Q.Fields[0].AsString ; Ferme(Q) ; Eplan.Text:=Plan ;
St:='' ;
for i:=1 to 6 do
  if RupClass[i] then St:=St+IntToStr(i)+'; ' ;
Edit1.Text:=Edit1.Text+St ;
END ;

Procedure TFDrupgene.TagDetag(Avec : Boolean) ;
Var  i : Integer ;
begin
for i:=1 to FListe.RowCount-1 do
    if Avec then FListe.Cells[FListe.ColCount-1,i]:='*'
            else FListe.Cells[FListe.ColCount-1,i]:='' ;

  FListe.Visible:=False ;
  FListe.Visible:=True ;
  FListe.SetFocus ;
  BDetag.Visible:=Avec ;
  BTag.Visible:=Not Avec ;
  CompteElemSelectionner ;
end;

procedure TFDrupgene.BdetagClick(Sender: TObject);
begin TagDetag(False) ; end;

procedure TFDrupgene.BTagClick(Sender: TObject);
begin TagDetag(True) ; end;

procedure TFDrupgene.FlisteMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if (ssCtrl in Shift) And (Button=mbLeft)then
    InverseSelection ;
end;

procedure TFDrupgene.BValiderClick(Sender: TObject);
Var i,j : Integer ;
    Trouver : Boolean ;
begin
Trouver:=False ; j:=0 ;
for i:=1 to FListe.RowCount-1 do
   if FListe.Cells[FListe.ColCount-1,i]='*' then BEGIN Trouver:=True ; Inc(j) ; END ;
if Trouver then
   BEGIN
    if MsgBox.Execute(0,'','')=mrYes then
       BEGIN
        BeginTrans ; InitMove(j,MsgBox.Mess[1]) ;
        ExecuteSql('Delete From RUPTURE Where RU_NATURERUPT="'+Nature+'" And '+
                   'RU_PLANRUPT="'+Plan+'" And RU_SOCIETE="'+V_PGI.CodeSociete+'"') ;
        for i:=1 to FListe.RowCount-1 do
          BEGIN
           if FListe.Cells[FListe.ColCount-1,i]='*' then
              BEGIN
              ExecuteSql('Insert Into RUPTURE (RU_NATURERUPT,RU_PLANRUPT,RU_CLASSE,'+
                          'RU_LIBELLECLASSE,RU_SOCIETE) '+
                          'Values ("'+Nature+'","'+Plan+'","'+Copy(FListe.Cells[0,i],1,17)+'",'+
                                  '"'+Copy(FListe.Cells[1,i],1,35)+'","'+V_PGI.CodeSociete+'")') ;
              MoveCur(False) ;
              END ;
          END ;
       FiniMove ; CommitTrans ;
       END ;
   END ;
Close ;
end;

Procedure DetailPlanRuptureGene(RupClass : TabBool ; LibClass : TabLib ; Nature,Plan : String) ;
var FDrupgene: TFDrupgene ;
BEGIN
if _Blocage(['nrCloture'],False,'nrAucun') then Exit ;
FDrupgene:=TFDrupgene.Create(Application) ;
  Try
   FDrupgene.RupClass:=RupClass ;
   FDrupgene.LibClass:=LibClass ;
   FDrupgene.Nature:=Nature ;
   FDrupgene.Plan:=Plan ;
   FDrupgene.ShowModal ;
  Finally
   FDrupgene.Free ;
  End ;
Screen.Cursor:=SyncrDefault ;
END ;

procedure TFDrupgene.BImprimerClick(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
  PrintDBGrid(Caption, FListe.ListeParam, '', '');
{$ELSE}
  PrintDBGrid(FListe,PTop,Caption,'') ;
{$ENDIF}
end;

procedure TFDrupgene.FormClose(Sender: TObject; var Action: TCloseAction);
begin
FListe.VidePile(True) ;
end;

procedure TFDrupgene.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFDrupgene.FormCreate(Sender: TObject);
begin
WMinX:=Width ; WMinY:=Height ;
end;

Procedure TFDrupgene.BrancheHelpContext ;
BEGIN
if Nature='RUG' then HelpContext:=1370300
                else HelpContext:=1380300 ;
END ;

procedure TFDrupgene.BAideClick(Sender: TObject);
begin CallHelpTopic(Self) ; end;

Function TFDrupgene.ListeVide : Boolean ;
BEGIN Result:=FListe.Cells[0,1]='' ; END ;

Procedure TFDrupgene.CompteElemSelectionner ;
var i,j : Integer ;
begin
  j:=0 ;
  if not ListeVide then
  begin
    for i:=1 to FListe.RowCount-1 do
       if FListe.Cells[FListe.ColCount-1,i]='*' then Inc(j) ;
  end;

  case j of
  0,1: begin
         Nb1.Caption  := IntTostr(j) ;
         Tex1.Caption := MsgBox.Mess[5] ;
       end;
  else begin
         Nb1.Caption  := IntTostr(j) ;
         Tex1.Caption := MsgBox.Mess[6] ;
       end;
  end;
end;

procedure TFDrupgene.FlisteKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if ((ssShift in Shift) and (Key=VK_DOWN)) or
      ((ssShift in Shift) And (Key=VK_UP)) then
    InverseSelection
  else
  begin
    if (Shift=[]) And (Key=VK_SPACE) then
    begin
      InverseSelection ;
      if (FListe.Row<FListe.RowCount-1) then
        FListe.Row:=FListe.Row+1 ;
    end;
  end;
end;

Procedure TFDrupgene.InverseSelection ;
begin
  if ListeVide then Exit ;
  if FListe.Cells[FListe.ColCount-1,FListe.Row] = '*' then
    FListe.Cells[FListe.ColCount-1,FListe.Row] := ''
  else
    FListe.Cells[FListe.ColCount-1,FListe.Row] := '*' ;
  CompteElemSelectionner ;
  FListe.Invalidate ;
end ;

Procedure TFDrupgene.GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
BEGIN
if FListe.Cells[FListe.ColCount-1,ARow]='*' then FListe.Canvas.Font.Style:=FListe.Canvas.Font.Style+[fsItalic]
                                            else FListe.Canvas.Font.Style:=FListe.Canvas.Font.Style-[fsItalic] ;
END ;

end.
