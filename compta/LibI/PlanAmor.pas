{***********UNITE*************************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 18/07/2003
Modifié le ... :   /  /
Description .. : - CA - 18/07/2003 - Suppression des champs euro
Mots clefs ... :
Suite .........: - MBO - 20/09/2005 - FQ 16734 - pb affichage plan amort si cession sur immo avec exceptionnel
Suite..........: - MBO - 30/09/2005 - FQ 16766 - prise en compte de l'exceptionnel ds dotation fiscale
Suite..........: - MBO - 24/10/2005 - crc 2002/10 - faire apparaitre except + depreciation dans l'exceptionnel
Suite..........: - MBO - 25/11/2005 - FQ 17082 - immo cédée le dernier jour de l'exercice :
                         exceptionnel cession non pris en compte car on faisait -1 sur date opération
Suite..........: - MBO - 30/11/2005 - FQ 17102 - impact de l'exceptionnel sur dotation fiscale
Suite..........: - MBO - 03/02/2006 - Ajout aux antérieurs eco et fiscal le montant antérieur dépréciation
                                      Modif de la fonction AfficheReprise
Suite..........: - MBO - 20/04/2006 - FQ 13047 - affichage dates d'exercice sur les lignes où l'amort n'est pas commencé
Suite..........: - MBO - 17/05/2006 - FQ 17569 - affichage des dates début d'amortissement
Suite..........: - MBO - 10/07/2006 - FQ 15274 - particularité affichage antérieur dérog en VARIABLE
Suite..........: - MBO - 13/07/2006 - FQ 18595 - particularité mode VAR en cas de cession pour calcul dérog
Suite..........: - TGA - 20/07/2006 - Ajout DPI affecté
Suite..........: - BTY - 28/07/2006 - PB compile en eaglclient sur DPI Affecté
Suite..........: - MBO - 25/08/2006 - correction de conseils de compilation
Suite..........: - MVG - 31/08/2006 - Modification du traitement des DPI suite CJ
Suite..........: - MBO -   /09/2006 - ajout d'un paramètre dans la f° AfficheDotationsIntermediaires pour gestion prime
Suite..........: - MBO - 18/10/2006 - modif du paramètre dans f° AfficheDotationsIntermediaires pour gestion subvention
Suite..........: - MBO - 28/05/2007 - modif de la dfm pour plafond déductibilité fiscale
Suite..........: - MBO - 10/07/2007 - ajout de TraduireMemoire sur les libellés liés au remplacement
Suite..........: - MBO - 31/12/2007 - fq 21309 - calcul d'une QP même si pas de plafond fiscal
Suite..........: - MBO - 10/09/2007 - fq 21405 - cession sur un exercice postérieur à l'exercice de fin d'amortissement
*****************************************************************}
unit PlanAmor;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, StdCtrls, HPanel, Buttons, HCtrls,ExtCtrls,HEnt1,Math,
  {$IFNDEF DBXPRESS}dbtables, HSysMenu, hmsgbox, HTB97, TntStdCtrls,
  TntGrids, TntExtCtrls{$ELSE}uDbxDataSet{$ENDIF},

  DBGrids, Spin, hmsgbox, ImEnt, MajTable, Hqry, DBCtrls,
  Mask, HTB97, UiUtil, ComCtrls, HDB, Menus, ImPlan, HRotLab, HSysMenu, ImEdCalc;

const
     CellDat : integer = 0;
     CellEco : integer = 1;
     CellFis : integer = 2;
     CellDer : integer = 3;
     CellRei : integer = 4;
     // ajout d'une colonne spéciale remplacement composant
     CellComp : integer = 5;

     LigRep : integer = 0;
     LigCum : integer = 1;
     LigVNC : integer = 2;
     LigCes : integer = 3;
     PremiereLigne : integer = 1;
type

  TPlanAmortissement = class(TForm)
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
    Label11: TLabel;
    DontCessionEco: TLabel;
    DontCessionFiscal: TLabel;
    DontExcepEco: TLabel;
    DontExcepFiscal: TLabel;
    VNCEco: TLabel;
    VNCFiscal: TLabel;
    CumulEco: TLabel;
    CumulFiscal: TLabel;
    CumulDerog: TLabel;
    CumulReint: TLabel;
    gAmortFiscal: TGroupBox;
    HLabel2: THLabel;
    lDureeFisc: TLabel;
    lMethodeFisc: TLabel;
    lMethodeFiscale: TLabel;
    HLabel3: THLabel;
    lTauxFisc: TLabel;
    lBaseFisc: TLabel;
    HLabel6: THLabel;
    gAmortEco: TGroupBox;
    HLabel4: THLabel;
    lDureeEco: TLabel;
    lMethodeEco: TLabel;
    Label3: TLabel;
    LabelTauxEco: THLabel;
    lTauxEco: TLabel;
    lBaseEco: TLabel;
    HLabel7: THLabel;
    gReintegrationFiscale: TGroupBox;
    Label4: TLabel;
    Label5: TLabel;
    MontantReintegration: TLabel;
    MontantDPI : TLabel;
    DPIaffecte : TLabel;
    MontantQuotePart: TLabel;
    bAnteriorite: TCheckBox;
    GroupBox1: TGroupBox;
    RepriseEco: TLabel;
    RepriseFiscal: TLabel;
    RepriseDerog: TLabel;
    HLabel1: THLabel;
    HLabel8: THLabel;
    HLabel9: THLabel;
    HLabel10: THLabel;
    CodeImmo: THLabel;
    DateAchat: THLabel;
    Label8: TLabel;
    ValeurHT: THLabel;
    HLabel11: THLabel;
    HMTrad: THSystemMenu;
    HLabel5: THLabel;
    HLabel12: THLabel;
    lDateDebEco: THLabel;
    lDateDebFis: THLabel;
    Lib_Reint: THLabel;
    RepriseReint: TLabel;
    DontCessionDerog: TLabel;
    DontCessionReint: TLabel;
    DontVR: TLabel;
    VR: TLabel;
    DontDeduc: TLabel;
    Deduc: TLabel;
    procedure GrilleAmortClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure GrilleAmortDblClick(Sender: TObject);
    procedure GrilleAmortMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure bAnterioriteClick(Sender: TObject);
    procedure bIntermediaireClick(Sender: TObject);
    procedure DessineGrille (Acol,ARow : LongInt ; Canvas : TCanvas;State : TGridDrawState) ;
    procedure GrilleAmortRowEnter(Sender: TObject; Ou: Integer;
      var Cancel: Boolean; Chg: Boolean);
    procedure BImprimerClick(Sender: TObject);
    procedure ToolbarButton973Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    PlanAmort: TPlanAmort;
    OldHint : string;
    DblCol : integer;
    dCumulEco : double;
    dCumulFisc : double;
    DateAffichageDot : TDateTime;

    { Déclarations privées }
    procedure InitGrilles;
    procedure InitDecPlan;
    procedure AfficheDotations;
    procedure AfficheMethodeAmort;
    procedure AfficheDotationCells(Col:integer;Lig:integer;Dot:double;DateExe: string);
    procedure AfficheReprise;
    procedure AfficheDotationReintegration(Col,Lig : integer;Reint,QPart:double);
  public
    { Déclarations publiques }
  end;

procedure AffichePlanAmortissement(Plan : TPlanAmort);

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

procedure AffichePlanAmortissement(Plan : TPlanAmort);
var
   FPlanAmort : TPlanAmortissement;
begin
  FPlanAmort := TPlanAmortissement.Create(Application);
  FPlanAmort.PlanAmort := Plan;
  try
    FPlanAmort.ShowModal;
  finally
    FPlanAmort.GrilleAmort.VidePile(true);
    FPlanAmort.free;
  end;
end;

procedure TPlanAmortissement.InitGrilles;
begin
  with GrilleAmort do
  begin
    PostDrawCell := DessineGrille;
    ColAligns[0]:=taCenter;
    ColAligns[CellEco]:=taRightJustify;
    ColAligns[CellFis]:=taRightJustify;
    ColAligns[CellDer]:=taRightJustify;
    ColAligns[CellRei]:=taRightJustify;
    // ajout mbo 04.06.07
    ColAligns[CellComp]:= taLeftJustify;
    GrilleAmort.ColFormats[5]:= '%4s';

    AlternateColor:=AltColors[V_PGI.NumAltCol];
  end;
end;

//pgr 07/2005 : FQ 16297 / gestion nombre de décimales
procedure TPlanAmortissement.InitDecPlan;
begin
  ValeurHT.DisplayFormat := StrFMask(V_PGI.OkDecV,'',True);
  lBaseEco.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
  lBaseFisc.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
  RepriseEco.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
  RepriseFiscal.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
  RepriseDerog.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
  RepriseReint.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
  MontantReintegration.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
  MontantDPI.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
  DontCessionEco.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
  DontCessionFiscal.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
  // fq 17512 ajout mbo
  DontCessionDerog.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
  DontCessionReint.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
  VR.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
  Deduc.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);

  DontExcepEco.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
  DontExcepFiscal.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
  CumulEco.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
  CumulFiscal.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
  CumulDerog.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
  CumulReint.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
  VNCEco.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
  VNCFiscal.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
end;

procedure TPlanAmortissement.DessineGrille (Acol,ARow : LongInt ; Canvas : TCanvas;State : TGridDrawState) ;
var R : TRect ;
begin
  if (not PlanAmort.Fiscal) and (ARow>0) and (ACol > 1) and (ACol<4) then
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
  //pour chantier fiscal   mbo 4.06.07 ajout du supérieur sur test acol
  if (PlanAmort.GestionFiscale= false) and (PlanAmort.ReintegrationFiscale = false)
     and (PlanAmort.Remplace = false) and (PlanAmort.EstRemplacee = false)
     and (ARow>0) and (ACol>=4) then
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

procedure TPlanAmortissement.FormShow(Sender: TObject);
var bAfficheFiscal : boolean;
begin
  InitGrilles;
  //pgr 07/2005 : FQ 16297 / gestion nombre de décimales
  InitDecPlan;

  CodeImmo.Caption := PlanAmort.CodeImmo;
  LibelleImmo.Caption := PlanAmort.LibelleImmo;

  //mbo pour gestion remplacement
  if (PlanAmort.GestionFiscale) then
     HLabel9.Caption := TraduireMemoire('fiscaux avec gestion fiscale')
  else
      HLabel9.Caption := TraduireMemoire('fiscaux sans gestion fiscale');

  // ajout mbo 04.06.07 les 2 controles sont visibles uniquement si composant remplacée
  if (PlanAmort.EstRemplacee = false) and (PlanAmort.Remplace = false) then
  begin
     DontVR.Visible :=false;
     VR.Visible := false;
     DontDeduc.visible := false;
     Deduc.Visible := false;
  end else
  begin
     if (PlanAmort.EstRemplacee = true) and (PlanAmort.Remplace = true) then
     begin
        DontVR.Visible :=true;
        VR.Visible := true;
        DontDeduc.visible := true;
        Deduc.Visible := true;
     end else
     begin
        // immo remplaçante
        if (PlanAmort.Remplace) then
        begin
          DontVR.Visible := false;
          DontDeduc.Visible := true;
          VR.visible := false;
          Deduc.Visible := true;
       end else
       begin
          // immo remplacée
          DontVR.Visible := true;
          DontDeduc.Visible := false;
          VR.visible := true;
          Deduc.Visible := false;
       end;
     end;
  end;


  // BTY 07/06 PB compile en eaglclient sur le Tquery
  // TGA 20/07/2006 lecture de Immomvtd pour cumul DPI sur immo
  {IF PlanAmort.CodeDPI = 'X' then
     Begin
       Q :=OpenSQL('SELECT * FROM IMMOMVTD WHERE IZ_IMMO="'+PlanAmort.CodeImmo+'"', FALSE) ;
       try
         While Not Q.Eof do
            begin
              Montant := Montant + Q.FindField('IZ_MONTANT').AsFloat;
              Q.Next ;
            End;
       finally
         Ferme(Q);
       End;
       MontantDPI.Caption:=StrFMontant(montant,15,V_PGI.OkDecV,'',True);
     End
  Else
     Begin
       DPIaffecte.Visible := False;
       MontantDPI.Visible := False;
     End; }
  //===============================================================================
  DPIaffecte.Visible := PlanAmort.DPI;
  MontantDPI.Visible := PlanAmort.DPI;
  if PlanAmort.DPI then
     MontantDPI.Caption := StrFMontant(PlanAmort.DPIAffectee, 15, V_PGI.OkDecV, '', True);
  //===============================================================================

  MontantReintegration.Caption:=StrFMontant(PlanAmort.ValReintegration,15,V_PGI.OkDecV,'',True);
  MontantQuotePart.Caption:=StrFMontant(PlanAmort.TauxQuotePart*100,15,V_PGI.OkDecV,'',True)+HM.Mess[7];
  DateAchat.Caption := DateToStr(PlanAmort.DateAchat);
  ValeurHT.Caption:=StrFMontant(PlanAmort.ValeurHT,15,V_PGI.OkDecV,'',True);
  AfficheMethodeAmort;
  DateAffichageDot:=VHImmo^.Encours.Fin;
  OldHint := GrilleAmort.Hint;
  AfficheDotations;
  AfficheReprise;
  bAfficheFiscal := PlanAmort.AmortFisc.Methode<>'';
  CumulFiscal.Visible := bAfficheFiscal;
  // fq 15274 - CumulDerog.Visible := bAfficheFiscal;
  VNCFiscal.Visible := bAfficheFiscal;
  DontCessionFiscal.Visible := bAfficheFiscal;
  DontExcepFiscal.Visible := bAfficheFiscal;
  //29/03/99 fait dans AfficheDotations
  //CumulReint.Visible := bAfficheFiscal;
  bAnteriorite.Visible := PlanAmort.ImmoCloture;
  if bAnteriorite.Visible then bAnterioriteClick(nil);
//  bIntermediaire.visible:=(GetCession(PlanAmort.CodeImmo,VHImmo^.Encours.Fin)<>'CES') ;
end;

procedure TPlanAmortissement.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if isInside(Self) then Action:=caFree ;
end;

procedure TPlanAmortissement.GrilleAmortDblClick(Sender: TObject);
var DotEco,Val1,Val2 : double;
begin
  if DblCol<>CellRei then exit;
  if GrilleAmort.Cells[CellEco,GrilleAmort.Row]='' then exit;
  DotEco:=Valeur(GrilleAmort.Cells[CellEco,GrilleAmort.Row]);
  Val1 := DotEco * PlanAmort.CoeffReintegration;
  Val2 := DotEco * PlanAmort.TauxQuotePart;
  OldHint := GrilleAmort.Hint;
  GrilleAmort.Hint := HM.Mess[8]+ ' ' +StrFMontant(Val1,15,V_PGI.OkDecV,'',True) +
                    ' # ' + HM.Mess[9]+ ' ' +StrFMontant(Val2,15,V_PGI.OkDecV,'',True);
  GrilleAmort.ShowHint:=true;
end;

procedure TPlanAmortissement.GrilleAmortMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var Lig : integer;
begin
  GrilleAmort.MouseToCell(X, Y, DblCol, Lig);
end;

// mbo 03.02.06 modif pour prise en compte antérieur dépréciation dans les cumul antérieurs eco et fiscal
procedure TPlanAmortissement.AfficheReprise;
var
   dRepriseEco: double;
   dRepriseFisc: double;

begin
  dRepriseEco  := PLanAmort.AmortEco.Reprise + PlanAmort.AmortEco.RepriseDep;
  dRepriseFisc := PLanAmort.AmortFisc.Reprise + PlanAmort.AmortFisc.RepriseDep;

  if dRepriseEco<>0 then RepriseEco.Caption := StrFMontant(dRepriseEco,15,V_PGI.OkDecV,'',True);
  if PlanAmort.AmortFisc.Methode<>'' then
  begin
    RepriseFiscal.Caption := StrFMontant(dRepriseFisc,15,V_PGI.OkDecV,'',True);

    // ajout mbo - fq 15274
    // chantier fiscal
    //if PlanAmort.AmortEco.Methode = 'VAR' then
    //begin
    //   if dRepriseFisc < dRepriseEco then
    //      RepriseDerog.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True)
    //      RepriseDerog.Caption :=
    //   else
    //      RepriseDerog.Caption := StrFMontant(dRepriseFisc-dRepriseEco,15,V_PGI.OkDecV,'',True);
    //end
    //else
       RepriseDerog.Caption := StrFMontant(PlanAmort.AmortDerog.Reprise,15,V_PGI.OkDecV,'',True);
       RepriseReint.Caption := StrFMontant(PlanAmort.AmortReint.Reprise,15,V_PGI.OkDecV,'',True);

  end;
end;

procedure TPlanAmortissement.AfficheDotationCells(Col:integer;Lig:integer;Dot:double;DateExe: string);
begin
 if Dot <> 0.00 then
 begin
    GrilleAmort.RowCount:=MaxIntValue([GrilleAmort.RowCount,Lig+1]);
    GrilleAmort.Cells[0,Lig]:=Format('%12s',[DateExe]);
    GrilleAmort.Cells[Col,Lig]:=StrFMontant(Dot,15,V_PGI.OkDecV,'',True);
 end
//29/03/99
 else
    // fq 13047- GrilleAmort.Cells[Col,Lig]:='';
 begin
    GrilleAmort.RowCount:=MaxIntValue([GrilleAmort.RowCount,Lig+1]);
    GrilleAmort.Cells[0,Lig]:=Format('%12s',[DateExe]);
 end
//29/03/99
end;

procedure TPlanAmortissement.GrilleAmortClick(Sender: TObject);
var
  dDotEco, dCesEco, dVNCEco, dExcEco, dGlobEco ,
  dCesFisc, dDotFisc, dVNCFisc, dExcFisc,
  dCumulDerog , dCumulReint, dValeurResidu : double;
  i : integer;
  DateLigne : TDateTime;
  DateRevisionNam : TDateTime;
  QuotePartVal, ReintegreVal : double;
begin
  //ReintegreVal := 0.00;
  // conseil de compil QuotePartVal := 0.00;
  dVNCFisc:=0.00;
  dVNCEco:=0.00;
  if OldHint<>'' then GrilleAmort.Hint:=OldHint;
  if not IsValidDate(GrilleAmort.Cells[CellDat,GrilleAmort.Row]) then exit;
  DateLigne:=StrToDate(GrilleAmort.Cells[CellDat,GrilleAmort.Row]);
  if GrilleAmort.Row=PremiereLigne then
  begin
    CumulEco.Caption   :=StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
    CumulFiscal.Caption:=StrfMontant(0.00,15,V_PGI.OkDecV,'',True);
    CumulDerog.Caption :=StrfMontant(0.00,15,V_PGI.OkDecV,'',True);
    CumulReint.Caption :=StrfMontant(0.00,15,V_PGI.OkDecV,'',True);

    if (PlanAmort.EstRemplacee = false) and (PlanAmort.Remplace = false) then
    begin
       DontVR.Visible :=false;
       VR.Visible := false;
       DontDeduc.visible := false;
       Deduc.Visible := false;
    end else
    begin
       if (PlanAmort.EstRemplacee = true) and (PlanAmort.Remplace = true) then
       begin
          DontVR.Visible :=true;
          VR.Visible := true;
          DontDeduc.visible := true;
          Deduc.Visible := true;
       end else
       begin
         // immo remplaçante
         if (PlanAmort.Remplace) then
         begin
            DontVR.Visible := false;
            DontDeduc.Visible := true;
            VR.visible := false;
            Deduc.Visible := true;
         end else
         begin
            // immo remplacée
            DontVR.Visible := true;
            DontDeduc.Visible := false;
            VR.visible := true;
            Deduc.Visible := false;
         end;
       end;
    end;
  end else
  begin
     DontVR.Visible :=false;
     VR.Visible := false;
     DontDeduc.visible := false;
     Deduc.Visible := false;
  end;

  // recupere le cumul des dotations effectuées jusqu'à ExerciceCourant-1
  PlanAmort.GetCumulsDotExercice(DateLigne-1,dCumulEco,dCumulFisc,false,true,false);
  dGlobEco:=PlanAmort.GetDotationGlobale(MethodeEco,DateLigne,dDotEco,dCesEco,dExcEco);

  { CA - 06/10/2003 - Inclusion de l'exceptionnel dans la sortie }
  if dCesEco<>0 then
  begin
    // mbo - fq 16734 - 20.09.2005 exeptionnel déja contenu dans dCesEco
    //DontCessionEco.Caption:=StrfMontant(dCesEco+dExcEco,15,V_PGI.OkDecV,'',True);
    DontCessionEco.Caption:=StrfMontant(dCesEco,15,V_PGI.OkDecV,'',True);
    // mbo - 24.10.05 dExcEco := PlanAmort.GetExcepExercice(DateLigne-1, DateRevisionNam, True ); // FQ 12370
    // dans dExcEco = exceptionnel + depreciation d'actif
    // mbo 25.11.06 fq 17082 - dExcEco := PlanAmort.GetExcepExercice(DateLigne-1, DateRevisionNam, True, True ); // FQ 12370
    dExcEco := PlanAmort.GetExcepExercice(DateLigne, DateRevisionNam, True, True );
  end else DontCessionEco.Caption:=StrfMontant(dCesEco,15,V_PGI.OkDecV,'',True);

  CumulEco.Caption:=StrfMontant(dGlobEco+dCumulEco-dCesEco,15,V_PGI.OkDecV,'',True); // Modif CA du 09/03/1999
  DontExcepEco.Caption:=StrfMontant(dExcEco,15,V_PGI.OkDecV,'',True);

  if PlanAmort.AmortEco.Base > 0 then
    //    dVNCEco := PlanAmort.AmortEco.Base-dDotEco-dCumulEco-dCesEco;
    dVNCEco := PlanAmort.AmortEco.Base-dDotEco-dCumulEco; // CA - 20/09/1999
  VNCEco.Caption:=StrfMontant(dVNCEco,15,V_PGI.OkDecV,'',True);

  if PlanAmort.Fiscal then
  begin
    PlanAmort.GetDotationGlobale(MethodeFisc,DateLigne,dDotFisc,dCesFisc,dExcFisc);
    // ajout mbo - fq 16766 - fq 17102
    // MBO 30/11/05 dExcFisc := dExcEco;
    DontCessionFiscal.Caption:=StrfMontant(dCesFisc,15,V_PGI.OkDecV,'',True);
    CumulFiscal.Caption:=StrfMontant(dDotFisc{+dCesFisc}+dCumulFisc,15,V_PGI.OkDecV,'',True); // Modif CA du 09/03/1999
    DontExcepFiscal.Caption:=StrfMontant(dExcFisc,15,V_PGI.OkDecV,'',True); // mbo - fq 16766

    if PlanAmort.AmortEco.Base > 0 then
    //      dVNCFisc := PlanAmort.AmortFisc.Base-dDotFisc-dCumulFisc-dCesFisc;
        dVNCFisc := PlanAmort.AmortFisc.Base-dDotFisc-dCumulFisc; // CA - 20/09/1999
    VNCFiscal.Caption:=StrfMontant(dVNCFisc,15,V_PGI.OkDecV,'',True);
  end;

  dCumulReint := 0.00;
  dCumulDerog := 0.00;

  // modif chantier fiscal
  { if PlanAmort.AmortEco.Methode = 'VAR' then
  begin
     If PlanAmort.AmortFisc.Reprise >= (PlanAmort.AmortEco.Reprise + PlanAmort.AmortEco.RepriseDep) then
        dCumulDerog := PlanAmort.AmortFIsc.Reprise - PlanAmort.amortEco.Reprise - PlanAMort.AmortEco.RepriseDep
     else
        dCumulReint := PlanAmort.amortEco.Reprise + PlanAMort.AmortEco.RepriseDep - PlanAmort.AmortFisc.Reprise;
  end else
  begin
     if PlanAmort.AmortFisc.Reprise >= PlanAmort.AmortEco.Reprise then
        dCumulDerog := PlanAmort.AmortFisc.Reprise - PlanAmort.AmortEco.Reprise
     else
        dCumulReint := PlanAmort.AmortEco.Reprise - PlanAmort.AmortFisc.Reprise;
  end; }


  // mbo -fq 15274 - ajout pour la gestion des antérieurs sur le plan variable
  { // chantier fiscal :
  if PlanAmort.AmortEco.Methode = 'VAR' then
  begin
     dCumulDerog := (PlanAmort.AmortFisc.Reprise -
                     PlanAmort.AmortEco.Reprise -
                     PlanAmort.AmortEco.RepriseDep);
     if dCumulDerog < 0 then dCumulDerog := 0.00;
  end;
  }

  // modif pour chantier fiscal
  dCumulDerog := PlanAmort.AmortDerog.Reprise;
  dCumulReint := PlanAmort.AmortReint.Reprise;

  if (dCesEco<> 0) or
     (dCesFisc <> 0) then
  begin

    // on traite une immo cédée  on traite le dérog à date de cession
    // par contre on traite la réintégration à date de cession 11.05.07
    PlanAmort.CumulDerogReint(DateLigne+1,dCumulDerog, dCumulReint);

    ReintegreVal := 0.0;
    QuotePartVal := 0.0;
    If (PlanAmort.DPIAffectee<>0.0) then
    begin
      QuotePartVal := 0.0;
      if PlanAmort.Fiscal then ReintegreVal := dCesFisc*(PlanAmort.CoeffReintegration) else ReintegreVal :=DCesEco*(PlanAmort.CoeffReintegration);
    end;

    If (PlanAmort.DPIAffectee=0.0) then
    begin
        if PlanAmort.Fiscal then
        begin
            if PlanAmort.ReintegrationFiscale then
            begin
               // fq 21309
               if PlanAmort.CoeffReintegration >0.0 then
               begin
                  ReintegreVal := DCesFisc*(1 - PlanAmort.CoeffReintegration);
                  QuotePartVal := DCesFisc*PlanAmort.CoeffReintegration*PlanAmort.TauxQuotePart;
               end else
                  QuotePartVal := DCesFisc*PlanAmort.TauxQuotePart;
            end;
        end else
        begin
            if PlanAmort.ReintegrationFiscale then
            begin
               // fq 21309
               if PlanAmort.CoeffReintegration > 0.0 then
               begin
                  ReintegreVal := DCesEco*(1 - PlanAmort.CoeffReintegration);
                  QuotePartVal := DCesEco*PlanAmort.CoeffReintegration*PlanAmort.TauxQuotePart;
               end else
                  QuotePartVal := DCesEco*PlanAmort.TauxQuotePart;
            end;
        end;
    end;

    dCumulReint := (dCumulReint* (-1)) + ReintegreVal + QuotePartVal;
    //ajout mbo le 10.05.07 on solde le dérogatoire et la réintégration
    dCumulDerog := dCumulDerog * (-1);

    // si l'immo a été remplacée on ajoute la valeur résiduelle comptable
    if (PlanAmort.EstRemplacee) then
    begin
       dValeurResidu := PlanAmort.GetValResiduelle(CodeImmo.Caption,PlanAmort.AmortEco,DateLigne+1, true);
       dCumulReint := dCumulReint + dValeurResidu;
       VR.Caption := StrfMontant(dValeurResidu,15,V_PGI.OkDecV,'',True);
    end;
    if (PlanAmort.Remplace) then
    begin
       Deduc.Caption := StrfMontant(PlanAmort.ValeurAchat,15,V_PGI.OkDecV,'',True);
       dCumulReint := (dCumulReint - PlanAmort.ValeurAchat);
    end;
    // modif mbo 01.06.07 dCumulReint := dCumulReint * (-1);
    //

    DontCessionDerog.Caption := StrfMontant(dCumulDerog,15,V_PGI.OkDecV,'',True);
    DontCessionReint.Caption := StrfMontant(dCumulReint,15,V_PGI.OkDecV,'',True);

    CumulDerog.Caption :=StrfMontant(0.00,15,V_PGI.OkDecV,'',True);
    CumulReint.Caption :=StrfMontant(0.00,15,V_PGI.OkDecV,'',True);


  end else
  begin
     DontCessionDerog.Visible := false;
     DontCessionReint.Visible := false;

     for i:=0 to GrilleAmort.Row-1 do
     begin
       // modif chantier fiscal
       //dCumulReint := dCumulReint+Valeur(GrilleAmort.Cells[4,i]);

       PlanAmort.CumulDerogReint(DateLigne+1,dCumulDerog, dCumulReint);

       ReintegreVal := 0.0;
       QuotePartVal := 0.0;
       If (PlanAmort.DPIAffectee<>0.0) then
       begin
         QuotePartVal := 0.0;
         if PlanAmort.Fiscal then ReintegreVal := Valeur(GrilleAmort.Cells[2,i+1])*(PlanAmort.CoeffReintegration) else ReintegreVal :=DCesEco*(PlanAmort.CoeffReintegration);
       end;

       If (PlanAmort.DPIAffectee=0.0) then
       begin
          if PlanAmort.Fiscal then
          begin
             if PlanAmort.ReintegrationFiscale then
             begin
               // ajout test sur coeffreintegration pour FQ 21309 - 31/08/2007
               if PlanAmort.CoeffReintegration > 0.0 then
               begin
                  ReintegreVal := Valeur(GrilleAmort.Cells[2,i+1])*(1 - PlanAmort.CoeffReintegration);
                  QuotePartVal := Valeur(GrilleAmort.Cells[2,i+1])*PlanAmort.CoeffReintegration*PlanAmort.TauxQuotePart;
               end else
                  QuotePartVal := Valeur(GrilleAmort.Cells[2,i+1])*PlanAmort.TauxQuotePart;
             end;
          end else
          begin
            if PlanAmort.ReintegrationFiscale then
            begin
               // ajout test sur coeffreintegration pour FQ 21309 - 31/08/2007
               if PlanAmort.CoeffReintegration > 0.0 then
               begin
                  ReintegreVal := Valeur(GrilleAmort.Cells[1,i+1])*(1 - PlanAmort.CoeffReintegration);
                  QuotePartVal := Valeur(GrilleAmort.Cells[1,i+1])*PlanAmort.CoeffReintegration*PlanAmort.TauxQuotePart;
               end else
                  QuotePartVal := Valeur(GrilleAmort.Cells[1,i+1])*PlanAmort.TauxQuotePart;
            end;
          end;
       end;

       dCumulReint := dCumulReint+ ReintegreVal + QuotePartVal ;
       //dCumulDerog := dCumulDerog + Valeur(GrilleAmort.Cells[CellDer,i+1]);
     end;

     // ajout mbo 01.06.07 particularité s/ 1ère ligne d'un composant on déduit le montant ht
     if (PlanAmort.Remplace) then
     begin
        dCumulReint := dCumulReint - PlanAmort.ValeurAchat;
        // ajout mbo 04.06.07
        Deduc.Caption := StrfMontant(PlanAmort.ValeurAchat,15,V_PGI.OkDecV,'',True);
     end;

     if (dCumulReint<>0.00) then
       CumulReint.Caption:=StrfMontant(dCumulReint,15,V_PGI.OkDecV,'',True);
     CumulReint.Visible := (dCumulReint<>0.00);

     if (dCumulDerog<>0.00) then
       CumulDerog.Caption:=StrfMontant(dCumulDerog,15,V_PGI.OkDecV,'',True);
     CumulDerog.Visible := (PlanAmort.AmortFisc.Methode<>'') and (dCumulDerog<>0.00);
  end;
end;

procedure TPlanAmortissement.AfficheMethodeAmort;
begin
  lMethodeEco.Caption := RechDom('TIMETHODEIMMO',PlanAmort.AmortEco.Methode,False);
  lDateDebEco.Caption:= DateToStr(PlanAmort.DateDebEco);    // ajout mbo fq 17569
  lDureeEco.Caption := IntToStr(PlanAmort.AmortEco.Duree) + HM.Mess[6];
  if PlanAmort.AmortEco.Methode = 'VAR' then
  begin
    lTauxEco.Caption := StrfMontant(PlanAmort.TotalUO,15,V_PGI.OkDecV,'',True);
    LabelTauxEco.Caption := TraduireMemoire('UO');
  end
  else lTauxEco.Caption := StrfMontant(PlanAmort.AmortEco.Taux*100,15,V_PGI.OkDecV,'',True) + HM.Mess[7];
  lBaseEco.Caption := StrfMontant(PlanAmort.AmortEco.Base,15,V_PGI.OkDecV,'',True);
  if PlanAmort.AmortFisc.Methode<>'' then
  begin
    lMethodeFiscale.Caption:=HM.Mess[11];
    lMethodeFisc.Caption := RechDom('TIMETHODEIMMO',PlanAmort.AmortFisc.Methode,False);
    lDateDebFis.Caption:= DateToStr(PlanAmort.DateDebFis);     // ajout mbo fq 17569
    lDureeFisc.Caption := IntToStr(PlanAmort.AmortFisc.Duree) + HM.Mess[6];
    lTauxFisc.Caption := StrfMontant(PlanAmort.AmortFisc.Taux*100,15,V_PGI.OkDecV,'',True) + HM.Mess[7];
    lBaseFisc.Caption := StrfMontant(PlanAmort.AmortFisc.Base,15,V_PGI.OkDecV,'',True);
  end
  else
  begin
    lMethodeFisc.Caption := '';
    lDateDebFis.Caption := '';   // ajout mbo fq 17569
    lDureeFisc.Caption := '';
    lTauxFisc.Caption := '';
    lBaseFisc.Caption := '';
    lMethodeFiscale.Caption:=HM.Mess[10];
  end;
end;

procedure TPlanAmortissement.AfficheDotations;  //utilisé pour les dotations intermédiaires
var
  DateExe : string;
  Lig,IndDot : integer;
  DotDerog,DotEco,DotFisc,ReintegreVal,QuotePartVal : double;
  DateDot : TDateTime;
  { Dérogatoire }
  CE, CF, AE, AF, EC, EA : double;
  i : integer;

  // ajout chantier fiscal
  ReintFisc , cumulReint : double;
  CumulDerog : double;

begin
  Lig:=1;
  GrilleAmort.VidePile(true);
  GrilleAmort.RowCount := 2;
  IndDot := 0;
  if not bAnteriorite.Checked then IndDot:=PlanAmort.DecalageAffDotations;

  CE := PlanAmort.AmortEco.Reprise + PlanAmort.AmortEco.RepriseCedee;
  CF := PlanAmort.AmortFisc.Reprise + PlanAmort.AmortFisc.RepriseCedee;
  AE := PlanAmort.AmortEco.Reprise + PlanAmort.AmortEco.RepriseCedee;
  AF := PlanAmort.AmortFisc.Reprise + PlanAmort.AmortFisc.RepriseCedee;
  // ajout mbo - fq 15274
  if PlanAmort.AmortEco.Methode = 'VAR' then
  begin
     // fq 18595 - prendre en compte les antÚrieurs en cas de cession
     CE := CE + PLanAmort.AmortEco.RepriseDep;
     AE := AE + PlanAmort.AmortEco.RepriseDep;
  end;

  CumulDerog := PlanAmort.AmortDerog.Reprise + PlanAmort.AmortDerog.RepriseCedee;
  CumulReint := PlanAmort.AmortReint.Reprise + PlanAmort.AmortReint.RepriseCedee;

  for i:=0 to IndDot - 1 do
  begin
    CE := CE + PlanAmort.AmortEco.TableauDot[i] +PlanAmort.AmortEco.TabCession[i];
    CF := CF + PlanAmort.AmortFisc.TableauDot[i] +PlanAmort.AmortFisc.TabCession[i];

    AE := CE + PlanAmort.AmortEco.TableauDot[i] +PlanAmort.AmortEco.TabCession[i];
    AF := CF + PlanAmort.AmortFisc.TableauDot[i] +PlanAmort.AmortFisc.TabCession[i];
  end;

  while PlanAmort.TableauDate[IndDot]<>iDate1900 do
  begin
    DateDot := PlanAmort.TableauDate[IndDot];
    DateExe := DateToStr(DateDot);
    { Dotation économique}
    DotEco:=0.00;
    if (PlanAmort.AmortEco.TableauDot[IndDot]<>0.00) or
       (PlanAmort.AmortEco.TabCession[IndDot]<>0.00) or
       (DateDot<>iDate1900) then
    begin
      DotEco:=PlanAmort.AmortEco.TableauDot[IndDot]+PlanAmort.AmortEco.TabCession[IndDot] ;
      { CA - 06/10/2003 - l'exceptionnel n'était pas pris en compte dans le calcul de la dotation en cas de sortie }
      // mbo - 20/09/2005 - fq 16734 - l'exceptionnel déjà pris dans DotEco
      //if PlanAmort.AmortEco.TabCession[IndDot] <>0 then
        //DotEco := DotEco+PlanAmort.GetExcepExercice(DateDot, DateOpe);
      AfficheDotationCells(CellEco,Lig,DotEco,DateExe);
    end;
    {Dotation fiscale}
    DotFisc:=0.00;
    if (PlanAmort.AmortFisc.TableauDot[IndDot] <> 0.00) or
       (PlanAmort.AmortFisc.TabCession[IndDot]<>0.00) or
       (DateDot<>iDate1900) then
    begin
      DotFisc:=PlanAmort.AmortFisc.TableauDot[IndDot]+PlanAmort.AmortFisc.TabCession[IndDot] ;
      AfficheDotationCells(CellFis,Lig,DotFisc,DateExe) ;
    end;

    //MVG 31/08/2006
    QuotePartVal := 0.0;
    ReintegreVal := 0.0;
    If (PlanAmort.DPIAffectee<>0.0) then
    begin
      QuotePartVal := 0.0;
      if PlanAmort.Fiscal then ReintegreVal := DotFisc*(PlanAmort.CoeffReintegration) else ReintegreVal :=DotEco*(PlanAmort.CoeffReintegration);
    end;
    //Fin MVG 31/08/2006
    If (PlanAmort.DPIAffectee=0.0) then
    begin
        if PlanAmort.Fiscal then
        begin
            if PlanAmort.ReintegrationFiscale then
            begin
               // ajout test coeffreintegration - fq 21309
               if PlanAmort.CoeffReintegration > 0.0 then
               begin
                  ReintegreVal := DotFisc*(1 - PlanAmort.CoeffReintegration);
                  QuotePartVal := DotFisc*PlanAmort.CoeffReintegration*PlanAmort.TauxQuotePart;
               end else
                  QuotePartVal := DotFisc*PlanAmort.TauxQuotePart;
            end;
        end else
        begin
            if PlanAmort.ReintegrationFiscale then
            begin
               //ajout test sur coeffreintegration - fq 21309
               if PlanAmort.CoeffReintegration > 0.0 then
               begin
                  ReintegreVal := DotEco*(1 - PlanAmort.CoeffReintegration);
                  QuotePartVal := DotEco*PlanAmort.CoeffReintegration*PlanAmort.TauxQuotePart;
               end else
                  QuotePartVal := DotEco*PlanAmort.TauxQuotePart;
            end;
        end;
    end;
    // déplacer apres le calcul derog : AfficheDotationReintegration(CellRei,Lig,ReintegreVal,QuotePartVal);

    // ajout mbo du test sur présence d'un plan fiscal pour effectuer ou non le calcul du dérogatoire 31.08.07
    if PlanAmort.Fiscal then
    begin
       { Dérogatoire }
       CE := CE + PlanAmort.AmortEco.TableauDot[IndDot] +PlanAmort.AmortEco.TabCession[IndDot];
       CF := CF + PlanAmort.AmortFisc.TableauDot[IndDot] +PlanAmort.AmortFisc.TabCession[IndDot];


       { FQ 15729 & 15728 - Ajout Arrondi }
       EC := Arrondi(CF-CE, V_PGI.OkDecV);
       EA := Arrondi(AF-AE,V_PGI.OkDecV);
       DotDerog := 0.0;

       // ajout pour chantier fiscal
       if PlanAmort.AmortEco.Methode = 'VAR' then
       begin

          if not (PlanAmort.GestionFiscale) then
          begin
               // on fait les calculs comme avant
               if (EC >= 0) then     // écart sur antérieurs
               begin
                 if EA < 0 then EA := 0;
                 if EC > EA then DotDerog := EC-EA
                 else if EA > EC then DotDerog := EC-EA;
               end else
                 if (EC < 0) then if (EA>0) then DotDerog := (-1)*EA;
          end else
          begin

             // ajout chantier fiscal : on solde le dérogatoire avant de générer la réintégration
             // mais pas l'inverse : c'est à dire qu'on se préoccupe pas du cumul réintégration avant de
             // générer du dérogatoire

             if (DotFisc > DotEco) then
             begin
               DotDerog := DotFisc - DotEco;
               CumulDerog := CumulDerog + DotDerog;
               ReintFisc := 0;
             end else
             begin
               if CumulDerog = 0 then
               begin
                  ReintFisc := (DotFisc - DotEco) * -1;
                  CumulReint := CumulReint + ReintFisc;
               end else
               begin
                  if CumulDerog > (DotFisc - DotEco)* -1 then
                  begin
                     DotDerog := (DotFisc - DotEco);  // on est en reprise de dérogatoire
                     CumulDerog := CumulDerog + DotDerog;
                  end else
                  begin
                     // on va solder le dérog et générer de la réintégration
                     DotDerog := CumulDerog * -1;
                     ReintFisc := (CumulDerog + DotFisc - DotEco)*-1;
                     CumulDerog := 0;
                     CumulReint := CumulReint + ReintFisc;
                  end;
               end;
             end;
          end;
       end else
       begin
          // le plan d'amortissement éco n'est pas variable
          if not (PlanAmort.GestionFiscale) then
          begin
               // on fait les calculs comme avant
             if (EC >= 0) then     // écart sur antérieurs
             begin
                 if EA < 0 then EA := 0;
                 if EC > EA then DotDerog := EC-EA
                 else if EA > EC then DotDerog := EC-EA;
             end else
                 if (EC < 0) then if (EA>0) then DotDerog := (-1)*EA;

          end else
          begin
                // ajout chantier fiscal : on génére une réintégration ou une déduction
                // qd c'est positif = réintégration - qd c'est négatif = déduction
                ReintFisc := DotEco - DotFisc;
                CumulReint := CumulReint + ReintFisc;
                DotDerog := 0;
          end;
       end;

       AE := CE;
       AF := CF;

       if (PlanAmort.AmortEco.TabCession[IndDot] <>0) or (PlanAmort.AmortFisc.TabCession[IndDot] <>0) then
       begin
         // mbo - FQ 21405 - 10.09.07 - attention aux cessions faites un exo postérieur à l'exo de fin d'amortissement
         if (PlanAmort.TableauDate[IndDot+1]=iDate1900) and (PlanAmort.CedeLe <= PlanAmort.TableauDate[IndDot]) then
         begin
           // l'immo est cédée  le dérog et la réint sont calculés à date de cession
           PlanAmort.CumulDerogReint(StrToDate(DateExe)+1, DotDerog, ReintFisc);

           ReintegreVal := ReintegreVal + ReintFisc;

           if PlanAmort.TableauDate[IndDot+1]=iDate1900 then
           begin
              If (PlanAmort.EstRemplacee) then
              begin
                 //dValeurResidu := PlanAmort.GetValResiduelle(CodeImmo.Caption,PlanAmort.AmortEco,StrToDate(DateExe)+1, true);
                 ReintegreVal := (ReintegreVal*(-1))+ PlanAmort.GetValResiduelle(CodeImmo.Caption,PlanAmort.AmortEco,StrToDate(DateExe)+1, true);
                 // l'immo a pu être une remplaçante avant d'être remplacée
                 if PlanAmort.Remplace then
                    ReintegreVal := ReintegreVal - PlanAmort.ValeurAchat;
                 GrilleAmort.Cells[CellComp,Lig]:= ' ***';
              end;
              // on a cédé un composant
              if (PlanAmort.Remplace) and not(PlanAmort.EstRemplacee) then
              begin
                 ReintegreVal := (ReintegreVal * (-1)) - PlanAmort.ValeurAchat;
                 GrilleAmort.Cells[CellComp,Lig]:= ' ***';
              end;
                 //VR.Caption := PlanAmort.GetValResiduelle(CodeImmo.Caption,PlanAmort.AmortEco,StrToDate(DateExe)+1, true);
              if not(PlanAmort.EstRemplacee) and not(PlanAmort.Remplace) then
              begin
                ReintegreVal := ReintegreVal * (-1);
                GrilleAmort.Cells[CellComp,Lig]:= '';
              end;

              DotDerog := DotDerog * (-1);
           end;
         end;
       end else
       begin
           ReintegreVal := ReintegreVal + ReintFisc;

           // particularité d'un composant remplaçant : on déduit la valeur HT de l'immo
           if (PlanAmort.Remplace) and (Lig = 1)  then
           begin
              ReintegreVal := ReintegreVal - PlanAmort.ValeurAchat;
              GrilleAmort.Cells[5,1]:= ' ***';
              //VR.Caption := StrfMontant(PlanAmort.ValeurHT,15,V_PGI.OkDecV,'',True);
           end;
       end;

       AfficheDotationCells(CellDer,Lig,DotDerog,DateExe);
    end;   //  fin du test si présence d'un plan fiscal

    if (PlanAmort.GestionFiscale = true) or (PlanAmort.ReintegrationFiscale = true)
        or (PlanAmort.Remplace = true) or (PlanAmort.EstRemplacee = true) then
       AfficheDotationReintegration(CellRei,Lig,ReintegreVal,QuotePartVal);

    // conseil de compil ReintegreVal := 0;
    ReintFisc := 0.0;

    Lig:=Lig+1;
    IndDot:=IndDot+1;
  end;
  GrilleAmortClick(Application);
end;

procedure TPlanAmortissement.AfficheDotationReintegration(Col,Lig : integer;Reint,QPart:double);
begin
  if Reint+QPart<>0 then
    GrilleAmort.Cells[Col,Lig]:=StrfMontant(Reint+QPart,15,V_PGI.OkDecV,'',True)
  else GrilleAmort.Cells[Col,Lig]:='';  //29/03/99
end;


procedure TPlanAmortissement.bAnterioriteClick(Sender: TObject);
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

procedure TPlanAmortissement.bIntermediaireClick(Sender: TObject);
begin
  AfficheDotationsIntermediaires (PlanAmort.CodeImmo,LibelleImmo.Caption, 0);  //mbo 18.10.06
end;

procedure TPlanAmortissement.GrilleAmortRowEnter(Sender: TObject;
  Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  if (Ou<PlanAmort.DecalageAffDotations+1) and (bAnteriorite.Checked)then
    Cancel := true;
end;

procedure TPlanAmortissement.BImprimerClick(Sender: TObject);
var titre: string;
begin
  titre := CodeImmo.Caption+' - '+LibelleImmo.caption;

  {$IFDEF EAGLCLIENT}
  PrintDBGrid (titre,GrilleAmort.name,'');
  {$ELSE}
  PrintDBGrid (GrilleAmort,Nil, titre,'');
  {$ENDIF}
end;

procedure TPlanAmortissement.ToolbarButton973Click(Sender: TObject);
begin
CallHelpTopic(Self);
end;

procedure TPlanAmortissement.FormCreate(Sender: TObject);
begin
{$IFDEF SERIE1}
HelpContext:=511020 ;
{$ELSE}
HelpContext:=2101100 ;
{$ENDIF}
end;

end.


