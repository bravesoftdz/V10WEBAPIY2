{***********UNITE*************************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 11/10/2006
Modifié le ... :   /  /
Description .. : - affichage du plan d'amortissement de la SUBVENTION d'équipement
Mots clefs ... :
Suite......... : - mbo - 24.10.2006 - si l'immo est cédée on ne veut pas voir apparaitre les dotations
                                      antérieures issues d'une clôture
Suite......... : - mbo - 20.12.2006 - fq 19355 - ajout de subvention ds le titre de l'impression                                      
Suite......... : - mbo - 26.04.2007 - modif de la dfm (suppression d'un caption sur un thpanel)
*****************************************************************}
unit PlanSBV;

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

  TfPlanSbv = class(TForm)
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
    LIB_DUREE: THLabel;
    lDuree: TLabel;
    lMethode: TLabel;
    Label3: TLabel;
    LabelTauxEco: THLabel;
    lTaux: TLabel;
    bAnteriorite: TCheckBox;
    CodeImmo: THLabel;
    DateAchat: THLabel;
    Label8: TLabel;
    ValeurHT: THLabel;
    HLabel11: THLabel;
    HMTrad: THSystemMenu;
    HLabel5: THLabel;
    lDateDeb: THLabel;
    Hmois: THLabel;
    MNTSBV: THLabel;
    li_montant: THLabel;
    gb_reprise: THGroupBox;
    HLabel8: THLabel;
    reprise: THLabel;
    lib_octroi: THLabel;
    lDateOctroi: THLabel;
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
    dCumulSBV : double;
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

procedure AffichePlanSbv(Plan : TPlanAmort);

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

procedure AffichePlanSbv(Plan : TPlanAmort);
var
   FPlanAmort : TfPlanSbv;
begin
  FPlanAmort := TfPlanSbv.Create(Application);
  FPlanAmort.PlanAmort := Plan;
  try
    FPlanAmort.ShowModal;
  finally
    FPlanAmort.GrilleAmort.VidePile(true);
    FPlanAmort.free;
  end;
end;

procedure TfPlanSbv.InitGrilles;
begin
  with GrilleAmort do
  begin
    PostDrawCell := DessineGrille;
    ColAligns[0]:=taCenter;
    ColAligns[CellEco]:=taRightJustify;
    AlternateColor:=AltColors[V_PGI.NumAltCol];
  end;
end;

procedure TfPlanSbv.InitDecPlan;
begin
  ValeurHT.DisplayFormat := StrFMask(V_PGI.OkDecV,'',True);
  Reprise.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
  DontCession.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
  Cumul.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
  VNC.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
  MNTSBV.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
end;

procedure TfPlanSbv.DessineGrille (Acol,ARow : LongInt ; Canvas : TCanvas;State : TGridDrawState) ;
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

procedure TfPlanSbv.FormShow(Sender: TObject);
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

procedure TfPlanSbv.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if isInside(Self) then Action:=caFree ;
end;


procedure TfPlanSbv.GrilleAmortMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var Lig : integer;
begin
  GrilleAmort.MouseToCell(X, Y, DblCol, Lig);
end;

procedure TfPlanSbv.AfficheReprise;
begin

  if (PlanAmort.AmortSBV.Reprise<>0) and
     (PlanAmort.AmortSBV.ReprisePrem = false) then
     Reprise.Caption := StrFMontant(PlanAmort.AmortSBV.Reprise,15,V_PGI.OkDecV,'',True);

end;


procedure TfPlanSbv.AfficheDotationCells(Col:integer;Lig:integer;Dot:double;DateExe: string);
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

procedure TfPlanSbv.GrilleAmortClick(Sender: TObject);
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
  PlanAmort.GetCumulSbv(DateLigne-1,dCumulSbv,false,true,false);


  dGlob:=PlanAmort.GetDotationGlobale(MethodeSBV,DateLigne,dDot,dCes,dExc);  // dotation exercice

  if dCes<>0 then
     DontCession.Caption:=StrfMontant(dCes,15,V_PGI.OkDecV,'',True)
  else
     DontCession.Caption:=StrfMontant(dCes,15,V_PGI.OkDecV,'',True);

  Cumul.Caption:=StrfMontant(dGlob+dCumulSBV-dCes,15,V_PGI.OkDecV,'',True);

  if PlanAmort.AmortSBV.Base > 0 then
      dVNC := PlanAmort.AmortSbv.Base-dDot-dCumulSBV;

  VNC.Caption:=StrfMontant(dVNC,15,V_PGI.OkDecV,'',True);

end;

procedure TfPlanSbv.AfficheMethodeAmort;
begin
  lMethode.Caption := RechDom('TIMETHODEIMMO',PlanAmort.SbvMethode,False);
  lDateDeb.Caption:= DateToStr(PlanAmort.SbvDateDeb);
  lDateOctroi.Caption := DateToStr(PlanAmort.SbvDateSBV);

  if PlanAmort.AmortEco.Methode = 'NAM' then
     lib_duree.Caption := 'Durée d''inaliénabilité'
  else
     lib_duree.Caption := 'Durée d''amortissement';

  lDuree.Caption := IntToStr(PlanAmort.SBVDuree);
  lTaux.Caption := StrfMontant(PlanAmort.SBVTaux,15,V_PGI.OkDecV,'',True) + HM.Mess[7];
  mntSbv.Caption := StrfMontant(PlanAmort.mntSBV,15,V_PGI.OkDecV,'',True);
end;

procedure TfPlanSBV.AfficheDotations;  //utilisé pour les dotations intermédiaires
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
  if planAmort.AmortSBV.ReprisePrem = false then
     CE := PlanAmort.AmortPri.Reprise
  else
     CE := 0.00;

  for i:=0 to IndDot - 1 do
  begin
    CE := CE + PlanAmort.AmortSbv.TableauDot[i] +PlanAmort.AmortSbv.TabCession[i];
  end;
  while PlanAmort.TableauDate[IndDot]<>iDate1900 do
  begin
    DateDot := PlanAmort.TableauDate[IndDot];
    DateExe := DateToStr(DateDot);
    { Dotation Subvention}
    if (PlanAmort.AmortSbv.TableauDot[IndDot]<>0.00) or
       (PlanAmort.AmortSbv.TabCession[IndDot]<>0.00) or
       ((DateDot >=PlanAmort.SBVDateDeb) and (DateDot<= PlanAmort.AmortSBV.DateFinAmort)) then

    begin
      DotEco:=PlanAmort.AmortSBV.TableauDot[IndDot]+PlanAmort.AmortSBV.TabCession[IndDot] ;

      // si l'immo est cédée on ne veut pas voir apparaitre les dotations des exercices précédents
      // modif mbo 24.10.06
      if (PlanAmort.AmortSBV.TableauDot[IndDot] = 0) and (PlanAmort.AmortSBV.TabCession[IndDot]<>0) then
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


procedure TfPlanSBV.bAnterioriteClick(Sender: TObject);
var
  OldCol : integer;
begin
  if (bAnteriorite.Checked) then OldCol:=GrilleAmort.RowCount+PlanAmort.DecalageAffDotations
                            else OldCol:=GrilleAmort.RowCount-PlanAmort.DecalageAffDotations;

  if OldCol<1 then exit ;
  AfficheDotations;
  OldCol:=Min(OldCol,GrilleAmort.RowCount-1);
  GrilleAmort.Row:=OldCol;
end;

procedure TfPlanSBV.bIntermediaireClick(Sender: TObject);
begin
  AfficheDotationsIntermediaires (PlanAmort.CodeImmo,LibelleImmo.Caption, 2);
end;

procedure TfPlanSBV.GrilleAmortRowEnter(Sender: TObject;
  Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  if (Ou<PlanAmort.DecalageAffDotations+1) and (bAnteriorite.Checked)then
    Cancel := true;
end;

procedure TfPlanSBV.BImprimerClick(Sender: TObject);
begin
  // fq 19355 - mbo - 20.12.02006 - préciser dans l'impression qu'il s'agit de la subvention

  {$IFDEF EAGLCLIENT}
  PrintDBGrid ('Subvention - ' + CodeImmo.Caption+' - '+LibelleImmo.caption,GrilleAmort.Name,'');
  {$ELSE}
  PrintDBGrid (GrilleAmort,Nil, 'Subvention - ' +CodeImmo.Caption+' - '+LibelleImmo.caption,'');
  {$ENDIF}
end;

procedure TfPlanSBV.ToolbarButton973Click(Sender: TObject);
begin
CallHelpTopic(Self);
end;

procedure TfPlanSBV.FormCreate(Sender: TObject);
begin

{$IFDEF SERIE1}
//HelpContext:=511020 ;
{$ELSE}
//HelpContext:=2101100 ;
{$ENDIF}

end;

end.


