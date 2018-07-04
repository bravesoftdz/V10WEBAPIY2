{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 04/01/2005
Modifié le ... :   /  /    
Description .. : Passage en eAGL
Mots clefs ... : 
*****************************************************************}
unit BudComp;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, Ent1,
  Hent1, hmsgbox, StdCtrls, Hctrls, ComCtrls, HRichEdt, HRichOLE, Buttons,
  ExtCtrls, SaisUtil, HSysMenu, Hcompte, LibChpLi;

Function SaisieBudComp(Li : TList ; Action : TActionFiche) : Boolean ;

type
  TFBudcomp = class(TForm)
    Panel1: TPanel;
    BValide: THBitBtn;
    BFerme: THBitBtn;
    BAide: THBitBtn;
    HM: THMsgBox;
    HmTrad: THSystemMenu;
    Pages: TPageControl;
    TS1: TTabSheet;
    GInt: TGroupBox;
    TBE_REFINTERNE: THLabel;
    TBE_LIBELLE: THLabel;
    BE_REFINTERNE: TEdit;
    BE_LIBELLE: TEdit;
    GQte: TGroupBox;
    TBE_QTE1: THLabel;
    TBE_QTE2: THLabel;
    TBE_QUALIFQTE1: THLabel;
    TBE_QUALIFQTE2: THLabel;
    TBE_AFFAIRE: THLabel;
    BE_QTE1: THNumEdit;
    BE_QTE2: THNumEdit;
    BE_QUALIFQTE1: THValComboBox;
    BE_QUALIFQTE2: THValComboBox;
    GBloc: TGroupBox;
    BE_BLOCNOTE: THRichEditOLE;
    TS2: TTabSheet;
    BE_AFFAIRE: TEdit;
    TBE_LIBRETEXTE1: THLabel;
    BE_LIBRETEXTE1: TEdit;
    TBE_LIBRETEXTE2: THLabel;
    BE_LIBRETEXTE2: TEdit;
    TBE_LIBRETEXTE3: THLabel;
    BE_LIBRETEXTE3: TEdit;
    TBE_LIBRETEXTE4: THLabel;
    BE_LIBRETEXTE4: TEdit;
    TBE_LIBRETEXTE5: THLabel;
    BE_LIBRETEXTE5: TEdit;
    TBE_TABLE0: THLabel;
    BE_TABLE0: THCpteEdit;
    TBE_TABLE1: THLabel;
    BE_TABLE1: THCpteEdit;
    TBE_TABLE2: THLabel;
    BE_TABLE2: THCpteEdit;
    TBE_TABLE3: THLabel;
    BE_TABLE3: THCpteEdit;
    procedure FormShow(Sender: TObject);
    procedure BE_REFINTERNEChange(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure BValideClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
    Li : TList ;
    Action : TActionFiche ;
    Modifier : Boolean ;
    Procedure LitLesEnreg ;
    Procedure EcritLesEnreg ;
    procedure EtudieZonesComp ;
  public
  end;


implementation

{$R *.DFM}

Function SaisieBudComp(Li : TList ; Action : TActionFiche) : Boolean ;
var FBudcomp : TFBudcomp ;
BEGIN
Result:=False ;
if Li=NIL then Exit ;
FBudcomp:=TFBudcomp.Create(Application) ;
  Try
   FBudcomp.Li:=Li ;
   FBudcomp.Action:=Action ;
   if FBudcomp.ShowModal=mrOk then Result:=True ;
  Finally
   FBudcomp.Free ;
  End ;
END ;

procedure TFBudcomp.EtudieZonesComp ;
Var i : integer ;
    C : TControl ;
    HH  : THLabel ;
    Nam,Nam1 : String ;
    Vis,Okok : boolean ;
BEGIN
Okok:=False ;
for i:=0 to TS2.ControlCount-1 do
    BEGIN
    C:=TControl(TS2.Controls[i]) ; if C is THLabel then Continue ;
    Nam:=C.Name ; Nam1:=Nam ; Vis:=True ;
    if PersoChamp('BE',Nam,Vis) then
       BEGIN
       HH:=THLabel(FindComponent('T'+Nam1)) ;
       if HH<>Nil then
          BEGIN
          HH.Caption:=Nam ;
          if Not Vis then C.Enabled:=False else Okok:=True ;
          if ((Not C.Enabled) and (Action<>taConsult)) then TEdit(C).Color:=clBtnFace ;
          END ;
       END else Okok:=True ;
    END ;
if Not Okok then Pages.Pages[1].TabVisible:=False ;
END ;

procedure TFBudcomp.FormShow(Sender: TObject);
begin
LitLesEnreg ; Modifier:=False ;
Case Action of
     taConsult : BEGIN Caption:=HM.Mess[2] ; FicheReadOnly(Self) ; END ;
     taModif   : Caption:=HM.Mess[1] ;
     taCreat   : Caption:=HM.Mess[0] ;
    End ;
EtudieZonesComp ;
end;

Procedure TFBudcomp.LitLesEnreg ;
Var O : TOBM ;
    C : TEdit ;
    k : integer ;
BEGIN
O:=TOBM(Li[0]) ;
if O=Nil then Exit ;
BE_REFINTERNE.Text:=O.GetMvt('BE_REFINTERNE') ; BE_LIBELLE.Text:=O.GetMvt('BE_LIBELLE') ;
BE_QTE1.Value:=O.GetMvt('BE_QTE1') ; BE_QTE2.Value:=O.GetMvt('BE_QTE2') ;
BE_QUALIFQTE1.Value:=O.GetMvt('BE_QUALIFQTE1') ; BE_QUALIFQTE2.Value:=O.GetMvt('BE_QUALIFQTE2') ;
StringsToRich(BE_BLOCNOTE,O.M) ;
BE_AFFAIRE.Text:=O.GetMvt('BE_AFFAIRE') ;
for k:=1 to 5 do
    BEGIN
    C:=TEdit(FindComponent('BE_LIBRETEXTE'+IntToStr(k))) ; if C=Nil then Continue ;
    C.Text:=O.GetMvt('BE_LIBRETEXTE'+IntToStr(k)) ;
    END ;
END ;

procedure TFBudcomp.BE_REFINTERNEChange(Sender: TObject);
begin Modifier:=True ; end;

procedure TFBudcomp.BFermeClick(Sender: TObject);
begin
ModalResult:=mrCancel ;
if Modifier then
   BEGIN
   if HM.Execute(3,'','')=mrYes then
      BEGIN
      EcritLesEnreg ; ModalResult:=mrOk ;
      END ;
   END ;
Close ;
end;

Procedure TFBudcomp.EcritLesEnreg ;
Var i,k : Integer ;
    O : TOBM ;
    C : TEdit ;
BEGIN
for i:=0 to Li.Count-1 do
  BEGIN
  O:=TOBM(Li[i]) ; if O=Nil then Continue ;
  O.PutMvt('BE_REFINTERNE',BE_REFINTERNE.Text)  ; O.PutMvt('BE_LIBELLE',BE_LIBELLE.Text) ;
  O.PutMvt('BE_QTE1',BE_QTE1.Value)             ; O.PutMvt('BE_QTE2',BE_QTE2.Value) ;
  O.PutMvt('BE_QUALIFQTE1',BE_QUALIFQTE1.Value) ; O.PutMvt('BE_QUALIFQTE2',BE_QUALIFQTE2.Value) ;
  RichToStrings(BE_BLOCNOTE,O.M) ;
  O.PutMvt('BE_AFFAIRE',BE_AFFAIRE.Text) ; 
  for k:=1 to 5 do
      BEGIN
      C:=TEdit(FindComponent('BE_LIBRETEXTE'+IntToStr(k))) ; if C=Nil then Continue ;
      O.PutMvt('BE_LIBRETEXTE'+IntToStr(k),C.Text) ;
      END ;
  END ;
END ;

procedure TFBudcomp.BValideClick(Sender: TObject);
begin
Modifier:=False ;
if Action<>taConsult then
   BEGIN
   EcritLesEnreg ;
   ModalResult:=mrOk ;
   END else
   BEGIN
   Close ;
   END ; 
end;

procedure TFBudcomp.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.
