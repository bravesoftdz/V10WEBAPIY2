unit ImSaiCoef;

interface

uses
 Classes,
 Hctrls,
 {$IFNDEF EAGLSERVER}
  HSysMenu, hmsgbox, Grids, HTB97, UiUtil, Controls, Forms, ParamDat,Hent1,uTob,hpanel,Windows,
 {$ENDIF}
 SysUtils ;

(*
{$IFDEF EAGLSERVER}
uses Classes, HSysMenu, hmsgbox, Grids, Hctrls, Controls, HTB97 ;
{$ELSE}
uses
  Windows,  SysUtils, Classes,  Controls, Forms,
  Grids, Hctrls, HTB97,utob,hpanel,UiUtil, HSysMenu, hmsgbox,ParamDat,Hent1;
*)

{$IFNDEF EAGLSERVER}
Procedure SaisieDesCoefficientsDegressifs((*Comment : TActionFiche*));
{$ENDIF}

Procedure InstalleLesCoefficientsDegressifs ( Nombase : string = '');

{$IFNDEF EAGLSERVER}
type
  TSaiCoefIm = class(TForm)
    Dock971: TDock97;
    HPB: TToolWindow97;
    BImprimer: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    Fliste: THGrid;
    BSupprimer: TToolbarButton97;
    BInserer: TToolbarButton97;
    HM1: THMsgBox;
    HMTrad: THSystemMenu;
    ToolbarButton971: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FlisteDblClick(Sender: TObject);
    procedure FlisteCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure FlisteGetEditMask(Sender: TObject; ACol, ARow: Integer; var Value: String);
    procedure BSupprimerClick(Sender: TObject);
    procedure BInsererClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ToolbarButton971Click(Sender: TObject);
    procedure FlisteKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FlisteCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure BAideClick(Sender: TObject);
  private
    { Déclarations privées }
    function IsLigneOK (Row : integer): boolean;
    //function IsLigneVide (Row : integer): boolean;
    procedure GoDetruitETRecreer ;
  public
    { Déclarations publiques }
  end;
{$ENDIF}

implementation

{$IFNDEF EAGLSERVER}

{$IFDEF eAGLClient}
uses utileAgl;
{$ELSE}
uses PrintDBG;
{$ENDIF}

{$ENDIF}

{$IFNDEF EAGLSERVER}
{$R *.DFM}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
Procedure InstalleLesCoefficientsDegressifs ( Nombase : string = '');
var stTable : string;
begin
  if NomBase = '' then stTable := 'CHOIXCOD'
  else stTable := NomBase+'.dbo.CHOIXCOD';
ExecuteSQL('DELETE FROM '+stTable+' WHERE CC_TYPE="ICD"') ;
ExecuteSQL('INSERT INTO '+stTable+' (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) VALUES ("ICD","001","01/01/1900;1.5;2;2.5","","")') ;
ExecuteSQL('INSERT INTO '+stTable+' (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) VALUES ("ICD","002","01/02/1996;2.5;3;3.5","","")') ;
ExecuteSQL('INSERT INTO '+stTable+' (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) VALUES ("ICD","003","01/02/1997;1.5;2;2.5","","")') ;
ExecuteSQL('INSERT INTO '+stTable+' (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) VALUES ("ICD","004","01/01/2001;1.25;1.75;2.25","","")') ;
end ;

{$IFNDEF EAGLSERVER}
Procedure SaisieDesCoefficientsDegressifs((*Comment : TActionFiche*));
var SaiCoefIm: TSaiCoefIm; PP : THPanel ;
begin
// CA - 25/10/2002 - Pas de coefficients dégressifs, ont les installe.
if not ExisteSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="ICD"') then
  InstalleLesCoefficientsDegressifs ;
SaiCoefIm:=TSaiCoefIm.Create(Application) ;
PP:=FindInsidePanel ;
if PP=Nil then
  begin
  try SaiCoefIm.ShowModal ; finally SaiCoefIm.Free ; end ;
  end
else
  begin
  InitInside(SaiCoefIm,PP) ;
  SaiCoefIm.Show ;
  end;
//Screen.Cursor:=SyncrDefault ;
end;


procedure TSaiCoefIm.FormShow(Sender: TObject);
var T1: Tob; i: integer ; St,wSt: string ;
begin
{$IFDEF SERIE1}
{$ELSE}
HelpContext:=2345010 ;
{$ENDIF}
Fliste.RowCount:=2 ;
for i:=0 to Fliste.ColCount-1 do
  begin
  FListe.Cells[i,1]:='' ;
  Fliste.ColAligns[i]:= taCenter;
  end ;
T1:=TOB.Create('UNKOWN',nil,-1) ;
T1.LoadDetailDB('CHOIXCOD','"ICD"','CC_CODE',nil,false) ;
for i:=0 to T1.Detail.Count-1 do
  begin
  St:=T1.Detail[i].GetValue('CC_LIBELLE') ;
  wSt:=READTOKENST(St) ; FListe.Cells[0,i+1]:=wSt ;
  FListe.Cells[1,i+1]:=DateToStr(iDate2099) ;
  wSt:=READTOKENST(St) ; FListe.Cells[2,i+1]:=wSt ;
  wSt:=READTOKENST(St) ; FListe.Cells[3,i+1]:=wSt ;
  wSt:=READTOKENST(St) ; FListe.Cells[4,i+1]:=wSt ;
  Fliste.RowCount:=Fliste.RowCount+1 ;
  if i>0 then FListe.Cells[1,i]:=DateToStr(StrToDate(FListe.Cells[0,i+1])-1) ;
  end ;
if Fliste.RowCount>2 then Fliste.RowCount:=Fliste.RowCount-1 ;
T1.free ;
Fliste.SetFocus ;
end;

procedure TSaiCoefIm.BValiderClick(Sender: TObject);
var i: integer ;
begin
if (HM1.Execute (1,Caption,'')<>mrYes) then exit ;
for i:=1 to FListe.RowCount-1 do if not IsLigneOk(i) then exit ;
if TRANSACTIONS(GoDetruitETRecreer,1)<>oeOk then HM1.Execute(9,caption,'')
                                   else ModalResult := mrYes ;
end;

procedure TSaiCoefIm.GoDetruitETRecreer ;
var T1,T2: tob; wSt: string ; i: integer ;
begin
ExecuteSql('DELETE FROM CHOIXCOD WHERE CC_TYPE="ICD"') ;
T1:=TOB.Create('UNKOWN',nil,-1) ;
for i:=1 to FListe.RowCount-1 do
  begin
  T2:=Tob.Create('CHOIXCOD',T1,-1) ;
  wSt:=FListe.Cells[0,i]+';'+FListe.Cells[2,i]
    +';'+FListe.Cells[3,i]+';'+FListe.Cells[4,i] ;
  T2.PutValue('CC_TYPE','ICD') ;
  T2.PutValue('CC_CODE',Format('%3.3d',[i])) ;
  T2.PutValue('CC_LIBELLE',wSt) ;
  end ;
T1.InsertOrUpdateDB ;
T1.free ;
end ;


(*function TSaiCoefIm.IsLigneVide (Row : integer): boolean;
begin
result:=((trim(FListe.Cells[1,Row])='/  /') or (trim(FListe.Cells[1,Row])=''))
  and (FListe.Cells[2,Row]='') and (FListe.Cells[3,Row]='') and (FListe.Cells[4,Row]='')
end ;*)

function TSaiCoefIm.IsLigneOK (Row : integer): boolean;
begin
result:=false ;
if (trim(FListe.Cells[0,Row])='//') or (FListe.Cells[2,Row]='')
      or (FListe.Cells[3,Row]='') or (FListe.Cells[4,Row]='') then begin  HM1.execute(4,caption,'') ; exit ; end ;
if not IsValidDate(FListe.Cells[0,Row]) then begin HM1.Execute(0,caption,'') ; exit ; end ;
//if (Row=1) and (StrToDate(FListe.Cells[1,Row])<StrToDate('01/01/2001')) then begin HM1.Execute(6,caption,'') ; exit ; end ;
if (Row<>1) and (StrToDate(FListe.Cells[0,Row-1])>=StrToDate(FListe.Cells[0,Row])) then begin HM1.Execute(7,caption,IntToStr(row)) ; exit ; end ;
result:=true ;
end ;

procedure TSaiCoefIm.BImprimerClick(Sender: TObject);
begin
{$IFDEF eAGLClient}
  PGIBox('PrintDBGrid','A FAIRE') ;
{$ELSE}
  PrintDBGrid (Fliste,Nil, Caption,'');
{$ENDIF eAGLCLient}
end;


procedure TSaiCoefIm.FlisteDblClick(Sender: TObject);
var wObj: ThEdit ; wC: char ;
begin
if Fliste.Col=0 then
  begin
  wObj:=ThEdit.Create(self) ;
  ParamDate(Self,wObj,wC) ;
  Fliste.Cells[0,FListe.Row]:=wObj.Text ;
  wObj.free ;
  end ;
end;

procedure TSaiCoefIm.FlisteCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var Cellule : string;
begin
Cellule:=Fliste.Cells[ACol,ARow];
case ACol of
  0 : if IsValidDate(Cellule)=false then begin HM1.Execute(0,caption,'') ; Cancel:=true end ;
  2..4 : if IsNumeric(Cellule)=false then begin HM1.Execute(8,caption,'') ; Cancel:=true end ;
  end ;
if (ACol=0) and (ARow>2) then Fliste.Cells[1,FListe.Row-1]:=DateToStr(StrToDate(Fliste.Cells[0,FListe.Row])-1) ;
if (ACol=3) and (ARow=Fliste.RowCount-1) then Fliste.RowCount:=Fliste.RowCount+1;
end;

procedure TSaiCoefIm.FlisteGetEditMask(Sender: TObject; ACol,
  ARow: Integer; var Value: String);
begin
  if (ACol=0) or (ACol=1) then Value := '##/##/####';
end;

procedure TSaiCoefIm.BSupprimerClick(Sender: TObject);
var i,j : integer;
begin
if Fliste.Row>1 then Fliste.Cells[1,FListe.Row-1]:=Fliste.Cells[1,FListe.Row] ;
for i:=Fliste.Row to Fliste.RowCount-1 do
  for j:= 0 to 5 do Fliste.Cells[j,i]:=Fliste.Cells[j,i+1];
if Fliste.RowCount>2 then
  Fliste.RowCount:=Fliste.RowCount-1
else
  for i:=0 to 5 do Fliste.Cells[i,1]:='';
end;

procedure TSaiCoefIm.BInsererClick(Sender: TObject);
var i,j : integer;
begin
if IsLigneOK (Fliste.Row) then
  begin
  Fliste.RowCount := Fliste.RowCount + 1;
  for i:=Fliste.RowCount downto Fliste.Row+1 do
    for j:=0 to 5 do
      begin
      Fliste.Cells[j,i]:= Fliste.Cells[j,i-1];
      Fliste.Cells[j,i-1] := '';
      end;
(*  if (FListe.Row=Fliste.RowCount-1) then Fliste.Cells[0,FListe.Row]:=DateToStr(iDate2099)
                                    else Fliste.Cells[1,FListe.Row]:=Fliste.Cells[0,FListe.Row+1] ;*)
  end;
end;

procedure TSaiCoefIm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if (ModalResult<>mrYes)and (HM1.Execute (5,Caption,'')<>mrYes) then Action:=caNone ;
end;

procedure TSaiCoefIm.ToolbarButton971Click(Sender: TObject);
begin
InstalleLesCoefficientsDegressifs ;
FormShow(nil) ;
end;

procedure TSaiCoefIm.FlisteKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
case Key of
  VK_DOWN : if (IsLigneOK(Fliste.Row)) and (Fliste.Row=Fliste.RowCount-1) then Fliste.RowCount:=Fliste.RowCount+1;
  end ;
end;

procedure TSaiCoefIm.FlisteCellEnter(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
if (ACol=0) then
  begin
  if (FListe.Row=Fliste.RowCount-1) then
    Fliste.Cells[1,FListe.Row]:=DateToStr(iDate2099) 
  else
    Fliste.Cells[1,FListe.Row]:=DateToStr(StrToDate(Fliste.Cells[0,FListe.Row+1])-1) ;
  end ;
end;

procedure TSaiCoefIm.BAideClick(Sender: TObject);
begin
  CallHelpTopic(Self);
end;
{$ENDIF}

end.
