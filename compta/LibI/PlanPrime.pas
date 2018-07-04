{***********UNITE*************************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 19/09/2006
Modifié le ... :   /  /
Description .. : - affichage du plan d'amortissement de la prime d'équipement
Mots clefs ... :
Suite ........ : - mbo - 18.10.2006 - modif du paramètre d'appel à f° AfficheDotationsIntermediaires
Suite......... : - mbo - 24.10.2006 - si l'immo est cédée on ne veut pas voir apparaitre les dotations
                                      antérieures issues d'une clôture
Suite......... : - mbo - 26.04.2007 - modif dfm (suppression caption sur un thpanel)
*****************************************************************}
unit PlanPrime;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, StdCtrls, HPanel, Buttons, HCtrls,ExtCtrls,HEnt1,DBTables, Db, Math,
  DBGrids, Spin, hmsgbox, ImEnt, MajTable, Hqry, DBCtrls, AmType,
  Mask, HTB97, UiUtil, ComCtrls, HDB, Menus, ImPlan, HRotLab, HSysMenu, ImEdCalc;

const
     CellDat : integer = 0;
     CellEco : integer = 1;

     LigRep : integer = 0;
     LigCum : integer = 1;
     LigVNC : integer = 2;
     LigCes : integer = 3;
     PremiereLigne : integer = 1;
type

  TfPlanPrime = class(TForm)
    HM: THMsgBox;
    HPanel3: THPanel;
    GrilleAmort: THGrid;
    HPanel2: THPanel;
    Label1: TLabel;
    LibelleImmo: THLabel;
    HPanel1: THPanel;
    ToolbarButton973: TToolbarButton97;
    BImprimer: TToolbarButton97;
    bFerme: TToolbarButton97;
    bIntermediaire: TToolbarButton97;
    HPanel4: THPanel;
    Label6: TLabel;
    Label7: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    DontCession: TLabel;
    VNC: TLabel;
    Cumul: TLabel;
    gAmortEco: TGroupBox;
    HLabel4: THLabel;
    lDuree: TLabel;
    lMethode: TLabel;
    Label3: TLabel;
    LabelTauxEco: THLabel;
    lTaux: TLabel;
    lBase: TLabel;
    HLabel7: THLabel;
    bAnteriorite: TCheckBox;
    GroupBox1: TGroupBox;
    Reprise: TLabel;
    HLabel8: THLabel;
    CodeImmo: THLabel;
    DateAchat: THLabel;
    Label8: TLabel;
    ValeurHT: THLabel;
    HLabel11: THLabel;
    HMTrad: THSystemMenu;
    HLabel5: THLabel;
    lDateDeb: THLabel;
    HPrime: THLabel;
    Prime: THLabel;
    Hmois: THLabel;
    procedure GrilleAmortClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure GrilleAmortMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure bAnterioriteClick(Sender: TObject);
    procedure bIntermediaireClick(Sender: TObject);
    procedure DessineGrille (Acol,ARow : LongInt ; Canvas : TCanvas;State : TGridDrawState ) ;
    procedure GrilleAmortRowEnter(Sender: TObject; Ou: Integer;
      var Cancel: Boolean; Chg: Boolean);
    procedure BImprimerClick(Sender: TObject);
    procedure ToolbarButton973Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);

  private
    PlanAmort: TPlanAmort;
    OldHint : string;
    DblCol : integer;
    dCumulPri : double;
    DateAffichageDot : TDateTime;
    { Déclarations privées }
    procedure InitGrilles;
    procedure InitDecPlan;
    procedure AfficheDotations;
    procedure AfficheMethodeAmort;
    procedure AfficheDotationCells(Col:integer;Lig:integer;Dot:double;DateExe: string);
    procedure AfficheReprise;
  public
    { Déclarations publiques }
  end;

procedure AffichePlanPrime(Plan : TPlanAmort);

implementation

uses  ImDotInt,
      ImOuPlan,
      {$IFDEF EAGLCLIENT}
      utileAGL
      {$ELSE}
      PrintDBG
      {$ENDIF}
      {$IFDEF SERIE1}
      {$ELSE}
      ,UtilPGI
      {$ENDIF};

{$R *.DFM}

procedure AffichePlanPrime(Plan : TPlanAmort);
var
   FPlanAmort : TfPlanPrime;
begin
  FPlanAmort := TfPlanPrime.Create(Application);
  FPlanAmort.PlanAmort := Plan;
  try
    FPlanAmort.ShowModal;
  finally
    FPlanAmort.GrilleAmort.VidePile(true);
    FPlanAmort.free;
  end;
end;

procedure TfPlanPrime.InitGrilles;
begin
  with GrilleAmort do
  begin
    PostDrawCell := DessineGrille;
    ColAligns[0]:=taCenter;
    ColAligns[CellEco]:=taRightJustify;
    AlternateColor:=AltColors[V_PGI.NumAltCol];
  end;
end;

procedure TfPlanPrime.InitDecPlan;
begin
  ValeurHT.DisplayFormat := StrFMask(V_PGI.OkDecV,'',True);
  lBase.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
  Reprise.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
  DontCession.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
  Cumul.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
  VNC.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
  Prime.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
end;

procedure TfPlanPrime.DessineGrille (Acol,ARow : LongInt ; Canvas : TCanvas;State : TGridDrawState) ;
var R : TRect ;
begin
  if (Acol = 1) and (GrilleAmort.CellValues[Acol, Arow] = '') then
  begin
    Canvas.Brush.Color := clBtnFace ;
    Canvas.Brush.Style := bsBDiagonal ;
    Canvas.Pen.Color   := clBtnFace ;
    Canvas.Pen.Mode    := pmCopy ;
    Canvas.Pen.Style   := psClear ;
    Canvas.Pen.Width   := 1 ;
    R:=GrilleAmort.CellRect(ACol,ARow) ;
    Canvas.Rectangle(R.Left,R.Top,R.Right+1,R.Bottom+1) ;
  end;

end;

procedure TfPlanPrime.FormShow(Sender: TObject);
begin
  InitGrilles;
  InitDecPlan;

  CodeImmo.Caption := PlanAmort.CodeImmo;
  LibelleImmo.Caption := PlanAmort.LibelleImmo;
  DateAchat.Caption := DateToStr(PlanAmort.DateAchat);
  ValeurHT.Caption:=StrFMontant(PlanAmort.ValeurHT,15,V_PGI.OkDecV,'',True);
  AfficheMethodeAmort;
  DateAffichageDot:=VHImmo^.Encours.Fin;
  OldHint := GrilleAmort.Hint;
  AfficheDotations;
  AfficheReprise;
  bAnteriorite.Visible := PlanAmort.ImmoCloture;
  if bAnteriorite.Visible then bAnterioriteClick(nil);
end;

procedure TfPlanPrime.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if isInside(Self) then Action:=caFree ;
end;


procedure TfPlanPrime.GrilleAmortMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var Lig : integer;
begin
  GrilleAmort.MouseToCell(X, Y, DblCol, Lig);
end;

procedure TfPlanPrime.AfficheReprise;
begin

  if PlanAmort.AmortPri.Reprise<>0 then Reprise.Caption :=
             StrFMontant(PlanAmort.AmortPri.Reprise,15,V_PGI.OkDecV,'',True);

end;

procedure TfPlanPrime.AfficheDotationCells(Col:integer;Lig:integer;Dot:double;DateExe: string);
begin
 if Dot <> 0.00 then
 begin
    GrilleAmort.RowCount:=MaxIntValue([GrilleAmort.RowCount,Lig+1]);
    GrilleAmort.Cells[0,Lig]:=Format('%12s',[DateExe]);
    GrilleAmort.Cells[Col,Lig]:=StrFMontant(Dot,15,V_PGI.OkDecV,'',True);
 end else
 begin
    GrilleAmort.RowCount:=MaxIntValue([GrilleAmort.RowCount,Lig+1]);
    GrilleAmort.Cells[0,Lig]:=Format('%12s',[DateExe]);
 end;
end;

procedure TfPlanPrime.GrilleAmortClick(Sender: TObject);
var
  dDot, dCes, dVNC, dExc, dGlob:double;
  DateLigne : TDateTime;
begin
  dVNC:=0.00;
  if OldHint<>'' then GrilleAmort.Hint:=OldHint;
  if not IsValidDate(GrilleAmort.Cells[CellDat,GrilleAmort.Row]) then exit;
  DateLigne:=StrToDate(GrilleAmort.Cells[CellDat,GrilleAmort.Row]);
  if GrilleAmort.Row=PremiereLigne then
  begin
    Cumul.Caption   :=StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
  end;

  // recupere le cumul des dotations effectuées jusqu'à ExerciceCourant-1  hors antérieurs saisis
  PlanAmort.GetCumulPri(DateLigne-1,dCumulPri,false,true,false);

  dGlob:=PlanAmort.GetDotationGlobale(MethodePri,DateLigne,dDot,dCes,dExc);  // dotation exercice

  if dCes<>0 then
     DontCession.Caption:=StrfMontant(dCes,15,V_PGI.OkDecV,'',True)
  else
     DontCession.Caption:=StrfMontant(dCes,15,V_PGI.OkDecV,'',True);

  Cumul.Caption:=StrfMontant(dGlob+dCumulPri-dCes,15,V_PGI.OkDecV,'',True);

  if PlanAmort.AmortPri.Base > 0 then
      dVNC := PlanAmort.AmortPri.Base-dDot-dCumulPri;

  VNC.Caption:=StrfMontant(dVNC,15,V_PGI.OkDecV,'',True);

end;

procedure TfPlanPrime.AfficheMethodeAmort;
begin
  lMethode.Caption := RechDom('TIMETHODEIMMO',PlanAmort.PriMethode,False);
  lDateDeb.Caption:= DateToStr(PlanAmort.PriDateDeb);

  if PlanAmort.AmortEco.Methode = 'NAM' then
     HLabel4.Caption := 'Durée d''inaliénabilité'
  else
     HLabel4.Caption := 'Durée d''amortissement';

  lDuree.Caption := IntToStr(PlanAmort.PRIDuree);
  lTaux.Caption := StrfMontant(PlanAmort.PRITaux,15,V_PGI.OkDecV,'',True) + HM.Mess[7];
  lBase.Caption := StrfMontant(PlanAmort.MNTPrime/2,15,V_PGI.OkDecV,'',True);
  Prime.Caption := StrfMontant(PlanAmort.MNTPrime,15,V_PGI.OkDecV,'',True);
end;

procedure TfPlanPrime.AfficheDotations;  //utilisé pour les dotations intermédiaires
var
  DateExe : string;
  Lig,IndDot : integer;
  DotEco, CE : double;
  DateDot : TDateTime;
  i : integer;
begin
  Lig:=1;
  GrilleAmort.VidePile(true);
  GrilleAmort.RowCount := 2;
  IndDot := 0;
  if not bAnteriorite.Checked then IndDot:=PlanAmort.DecalageAffDotations;
  CE := PlanAmort.AmortPri.Reprise;

  for i:=0 to IndDot - 1 do
  begin
    CE := CE + PlanAmort.AmortPri.TableauDot[i] +PlanAmort.AmortPri.TabCession[i];
  end;

  while PlanAmort.TableauDate[IndDot]<>iDate1900 do
  begin
    DateDot := PlanAmort.TableauDate[IndDot];
    DateExe := DateToStr(DateDot);
    { Dotation économique}
    if (PlanAmort.AmortPri.TableauDot[IndDot]<>0.00) or
       (PlanAmort.AmortPri.TabCession[IndDot]<>0.00) or
       ((DateDot >=PlanAmort.PRIDateDeb) and (DateDot<= PlanAmort.AmortPri.DateFinAmort)) then
       //(DateDot<>iDate1900) then
    begin
      DotEco:=PlanAmort.AmortPri.TableauDot[IndDot]+PlanAmort.AmortPri.TabCession[IndDot] ;

      // si l'immo est cédée on ne veut pas voir apparaitre les dotations des exercices précédents
      // modif mbo 24.10.06
      if (PlanAmort.AmortPri.TableauDot[IndDot] = 0) and (PlanAmort.AmortPri.TabCession[IndDot]<>0) then
      begin
        if IndDot < MAX_LIGNEDOT then
        begin
          if (PlanAmort.TableauDate[IndDot+1] <> iDate1900) then
              DotEco := 0;
        end;
      end;
      AfficheDotationCells(CellEco,Lig,DotEco,DateExe);
    end;

    Lig:=Lig+1;
    IndDot:=IndDot+1;
  end;
  GrilleAmortClick(Application);
end;


procedure TfPlanPrime.bAnterioriteClick(Sender: TObject);
var
  OldCol : integer;
begin
  if (bAnteriorite.Checked) then OldCol:=GrilleAmort.RowCount+PlanAmort.DecalageAffDotations
                            else OldCol:=GrilleAmort.RowCount-PlanAmort.DecalageAffDotations;
//  if OldCol<2 then exit ; //YCP 19/09/01
  if OldCol<1 then exit ; //YCP 19/09/01
  AfficheDotations;
  OldCol:=Min(OldCol,GrilleAmort.RowCount-1);
  GrilleAmort.Row:=OldCol;
end;

procedure TfPlanPrime.bIntermediaireClick(Sender: TObject);
begin
  AfficheDotationsIntermediaires (PlanAmort.CodeImmo,LibelleImmo.Caption, 1); // mbo 18.10.06
end;

procedure TfPlanPrime.GrilleAmortRowEnter(Sender: TObject;
  Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  if (Ou<PlanAmort.DecalageAffDotations+1) and (bAnteriorite.Checked)then
    Cancel := true;
end;

procedure TfPlanPrime.BImprimerClick(Sender: TObject);
begin
  {$IFDEF EAGLCLIENT}
  PrintDBGrid (CodeImmo.Caption+' - '+LibelleImmo.caption,GrilleAmort.Name,'');
  {$ELSE}
  PrintDBGrid (GrilleAmort,Nil, CodeImmo.Caption+' - '+LibelleImmo.caption,'');
  {$ENDIF}
end;

procedure TfPlanPrime.ToolbarButton973Click(Sender: TObject);
begin
CallHelpTopic(Self);
end;

procedure TfPlanPrime.FormCreate(Sender: TObject);
begin

{$IFDEF SERIE1}
//HelpContext:=511020 ;
{$ELSE}
//HelpContext:=2101100 ;
{$ENDIF}

end;

end.


