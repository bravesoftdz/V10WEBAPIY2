unit Revision;

interface

uses
Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
StdCtrls, Buttons, Hctrls, ComCtrls, HQuickRP, ExtCtrls, DB, DBTables,
hmsgbox, Ent1, HQry, HEnt1, Mask, Hcompte, CpteUtil, ParamDat,
SaisUtil, Filtre, Menus, HSysMenu, Paramsoc
  ,HPanel, UIUtil, HTB97 // MODIF PACK AVANCE pour gestion mode inside
;

procedure LanceRevision ;

type
  TRevis = class(TForm)
    Panel1: TPanel;
    TDateRev: THLabel;
    FDateRev: TMaskEdit;
    Confirmation: THMsgBox;
    EtatRevision: TRadioGroup;
    HMTrad: THSystemMenu;
    Dock: TDock97;
    HPB: TToolWindow97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    procedure BValiderClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure EtatRevisionClick(Sender: TObject);
    procedure FDateRevEnter(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Uses UtilPgi ;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 09/09/2003
Modifié le ... :   /  /    
Description .. : 
Suite ........ : 09/09/2003, SBO : MODIF PACK AVANCE pour gestion 
Suite ........ : mode inside
Mots clefs ... : 
*****************************************************************}
procedure LanceRevision ;
var Revis: TRevis;
    PP : THPanel ;
BEGIN
  if Not _BlocageMonoPoste(True) then Exit ;

  Revis:=TRevis.Create(Application) ;

  PP:=FindInsidePanel ;
  if PP=Nil then
    begin
    try
     Revis.ShowModal ;
     finally
     Revis.Free ;
     end ;
    end
  else
    begin
    InitInside(Revis,PP) ;
    Revis.Show ;
    end ;

  _DeblocageMonoPoste(True) ;
  Screen.Cursor:=SyncrDefault ;
END ;

procedure TRevis.BValiderClick(Sender: TObject);
var DateRevision : TDateTime ;
    Reponse      : Integer ;
begin
DateRevision:=0 ;
Reponse:=Confirmation.Execute(EtatRevision.ItemIndex,'','') ;
if Reponse=mrYes then
   BEGIN
   BValider.Enabled:=False ;
   Case EtatRevision.ItemIndex of
     0 : BEGIN DateRevision:=StrToDate(FDateRev.Text) ; END ;
     1 : BEGIN DateRevision:=IDate1900 ; FDateRev.Text:=DateToStr(DateRevision) ; END ;
    end ;
   If DateRevision<>0 Then
      BEGIN
{$IFDEF SPEC302}
      ExecuteSQL('UPDATE SOCIETE SET SO_DATEREVISION="'+USDateTime(DateRevision)+'" WHERE SO_SOCIETE="'+V_PGI.CodeSociete+'"') ;
{$ELSE}
      SetParamSoc('SO_DATEREVISION',DateRevision) ;
{$ENDIF}
      VH^.DateRevision:=DateRevision ;
      END ;
   END ;
end;

procedure TRevis.FormShow(Sender: TObject);
Var StDate : string ;
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
HorzScrollBar.Range:=0 ;HorzScrollBar.Visible:=FALSE ;
VertScrollBar.Range:=0 ;VertScrollBar.Visible:=FALSE ;
//ClientHeight:=HPB.Top+HPB.Height ; ClientWidth:=HPB.Left+HPB.Width ;
StDate:=DateToStr(VH^.DateRevision) ; FDateRev.Text:=StDate ;
if VH^.DateRevision=IDate1900 then
   BEGIN
   EtatRevision.ItemIndex:=1 ;
   TDateRev.Visible:=False ;
   FDateRev.Visible:=False ;
   END else
   BEGIN
   EtatRevision.ItemIndex:=0 ;
   TDateRev.Visible:=True ;
   FDateRev.Visible:=True ;
   FDateRev.Text:=StDate ;
   END ;
BValider.Enabled:=False ;
end;

procedure TRevis.EtatRevisionClick(Sender: TObject);
begin
BValider.Enabled:=True ;
case EtatRevision.ItemIndex of
  0 : BEGIN TDateRev.Visible:=True ; FDateRev.Visible:=True ; END ;
  1 : BEGIN TDateRev.Visible:=False ; FDateRev.Visible:=False ; END;
 end ;
end;

procedure TRevis.FDateRevEnter(Sender: TObject);
begin
BValider.Enabled:=True ;
end;

procedure TRevis.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ; 
end;

procedure TRevis.BFermeClick(Sender: TObject);
begin
  Close;
  if IsInside(Self) then
  begin
      CloseInsidePanel(Self) ;
  end;
end;

end.
