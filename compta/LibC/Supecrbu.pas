{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 11/01/2005
Modifié le ... :   /  /    
Description .. : Remplacé en eAGL par SUPECRBU_TOF.PAS
Mots clefs ... : 
*****************************************************************}
unit Supecrbu;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, HSysMenu, Menus, DB, DBTables, Hqry, StdCtrls, Grids, DBGrids, HDB,
  ComCtrls, HRichEdt, Hctrls, ExtCtrls, Buttons, Hcompte, Mask, hmsgbox,
  Ent1, Hent1, SaisBud, SaisUtil, HTB97, ColMemo, HPanel, UiUtil, HStatus,
  HRichOLE ;

Procedure DetruitBudgets ;

type
  TFSupecrbu = class(TFMul)
    HM: THMsgBox;
    BE_BUDJAL: THValComboBox;
    TBE_BUDJAL: TLabel;
    TBE_NUMEROPIECE: TLabel;
    BE_NUMEROPIECE: THCritMaskEdit;
    BE_NUMEROPIECE_: THCritMaskEdit;
    TBE_NUMEROPIECE_: TLabel;
    BE_DATECOMPTABLE: THCritMaskEdit;
    BE_DATECOMPTABLE_: THCritMaskEdit;
    TBE_DATECOMPTABLE2: THLabel;
    TBE_DATECOMPTABLE: THLabel;
    TBE_EXERCICE: THLabel;
    BE_EXERCICE: THValComboBox;
    BE_AXE: THValComboBox;
    TBE_AXE: TLabel;
    TBE_NATURE: TLabel;
    BE_NATURE: THValComboBox;
    TBE_QUALIFPIECE: THLabel;
    BE_QUALIFPIECE: THValComboBox;
    TBE_UTILISATEUR: THLabel;
    BE_UTILISATEUR: THValComboBox;
    TBE_ETABLISSEMENT: TLabel;
    BE_ETABLISSEMENT: THValComboBox;
    BE_REFINTERNE: TEdit;
    TBE_REFINTERNE: THLabel;
    TBE_GENERAL: THLabel;
    BE_BUDGENE: THCpteEdit;
    BE_VALIDE: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure FListeDblClick(Sender: TObject); override;
    procedure BOuvrirClick(Sender: TObject); override;
    procedure QBeforeOpen(DataSet: TDataSet);
  private
    procedure DetruitLaPiece ;
  public
  end;


implementation

{$R *.DFM}

Procedure DetruitBudgets ;
var FSupecrbu : TFSupecrbu ;
    PP : THPanel ;
BEGIN
FSupecrbu:=TFSupecrbu.Create(Application) ;
FSupecrbu.Q.Manuel:=TRUE ;
FSupecrbu.FNomFiltre:='SUPPRBUDE' ;
FSupecrbu.Q.Liste:='SUPPRBUDE' ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FSupecrbu.ShowModal ;
    finally
     FSupecrbu.Free ;
    end;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FSupecrbu,PP) ;
   FSupecrbu.Show ;
   END ;
END ;

procedure TFSupecrbu.FormShow(Sender: TObject);
begin
BE_DATECOMPTABLE.text:=StDate1900 ; BE_DATECOMPTABLE_.text:=StDate2099 ;
PositionneEtabUser(BE_ETABLISSEMENT) ;
  inherited;
end;

procedure TFSupecrbu.FListeDblClick(Sender: TObject);
begin
  inherited;
if (Q.Eof) And (Q.Bof) then Exit ;
TrouveEtLanceSaisBud(Q,taConsult) ;
end;


procedure TFSupecrbu.DetruitLaPiece ;
BEGIN
if ExecuteSql('Delete From BUDECR Where BE_BUDJAL="'+Q.FindField('BE_BUDJAL').AsString+'" '+
              'And BE_NATUREBUD="'+Q.FindField('BE_NATUREBUD').AsString+'" '+
              'And BE_NUMEROPIECE='+IntToStr(Q.FindField('BE_NUMEROPIECE').AsInteger)+' '+
              'And BE_QUALIFPIECE="'+Q.FindField('BE_QUALIFPIECE').AsString+'" '+
              'And BE_VALIDE="-"')<=0 then V_PGI.IoError:=oeUnknown ;
END ;

procedure TFSupecrbu.BOuvrirClick(Sender: TObject);
Var i,NbD : integer ;
begin
//  inherited;
if (Q.Eof) And (Q.Bof) then Exit ;
if HM.Execute(0,'','')<>mrYes then Exit ;
NbD:=FListe.NbSelected ;
if NbD<=0 then BEGIN HM.Execute(1,caption,'') ; Exit ; END ;
Application.ProcessMessages ;
InitMove(NbD,'') ;
for i:=0 to NbD-1 do
    BEGIN
    FListe.GotoLeBookMark(i) ; MoveCur(FALSE) ;
    if Transactions(DetruitLaPiece,3)<>oeOK then BEGIN MessageAlerte(HM.Mess[2]) ; FiniMove ; Exit ; END ;
    END ;
FListe.ClearSelected ; FiniMove ;
BChercheClick(Nil) ;
END ;


procedure TFSupecrbu.QBeforeOpen(DataSet: TDataSet);
begin
  inherited;
  Q.SQL.Text := ModifieRequeteGrille(Q.SQL.Text); {FP FQ16049}
end;

end.
