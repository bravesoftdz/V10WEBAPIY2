{***********UNITE*************************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 07/11/2005
Modifié le ... : 07/11/2005
Description .. : BTY - FQ 16999 Le cumul antérieur ECO est faux pour toute
Suite ........ : date de calcul autre que date fin d'exercice
Suite......... : 31/03/2006 - MBO - FQ 15380 - dérogatoire faux si plan variable (ne tenait pas compte de l'antérieur dérog)
Suite......... : 02/06/2006 - MBO - FQ 15848 - ajout de l'aide + modif 19.06.06 - appel bouton faux sur v° antérieures
Suite......... : 12/07/2006 - MVG - Correction conseil de compilation
Suite......... : 20/09/2006 - MBO - ajout du paramètre pri en entrée pour affichage des dot prime
Suite......... : 18/10/2006 - MBO - Modif du paramètre en entrée et ajout gestion subvention
Suite......... : 08/03/2007 - MBO - Ajout d'une ligne réintégration/Déduction fq 17512
Suite......... : 31/05/2007 - MBO - modif double mouette en simple mouette
Suite......... :    06/2007 - MBO - gestion des composants de 2ème catégorie
Suite......... : 10/07/2007 - MBO - ajout de TraduireMemoire sur les libellés non gérés par la dfm
Suite......... : 02/10/2007 - MBO - fq 21754 - ajout d'un paramètre ds appel planInfo.calcul  (appel pour calcul cession)
Suite......... : 29/10/2007 - MBO - fq 21754 - ajout d'un paramètre ds appel planInfo.calcul (prise en cpte du jour de cession)
Mots clefs ... :
*****************************************************************}
unit ImDotInt;

interface

uses
  Windows,
  Messages,
  SysUtils,
  Classes,
  Graphics,
  Controls,
  Forms,
  Dialogs,
  Grids,
  Hctrls,
  StdCtrls,
  Mask,
  HTB97,
  ExtCtrls,
  HPanel,
  HSysMenu,
  hmsgbox,
  PlanAmor,
  HEnt1,
  {$IFDEF EAGLCLIENT}
  utileAGL,
  {$ELSE}
  PrintDBG,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  {$ENDIF}
  ImPlan,
  ParamDat,
  ImEnt;

const
  COL_CUMANT = 1;
  COL_DOTCALC = 2;
  COL_CUMAMORT = 3;
  COL_VNC = 4;
  LIG_ECO = 1;
  LIG_FISC = 2;

type
  TFDotInt = class(TForm)
    HPanel2: THPanel;
    HLabel1: THLabel;
    fDateCalcul: THCritMaskEdit;
    fGrid: THGrid;
    HM: THMsgBox;
    HMTrad: THSystemMenu;
    Dock971: TDock97;
    PBouton: TToolWindow97;
    ToolbarButton971: TToolbarButton97;
    BAPPLIQUER: TToolbarButton97;
    ToolbarButton974: TToolbarButton97;
    BoutonAide: TToolbarButton97;
    pAmortDerog: TPanel;
    tAmortDerog: THLabel;
    AmortDerog: THLabel;
    tReintDeduc: THLabel;
    ReintDeduc: THLabel;
    typederog: THLabel;
    typereint: THLabel;
    InfoReint: THLabel;
    InfoDeduc: THLabel;
    procedure FormShow(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure fDateCalculKeyPress(Sender: TObject; var Key: Char);
    procedure PostDrawCell(Acol,ARow : LongInt ; Canvas : TCanvas;State :TGridDrawState);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BValiderClick(Sender: TObject);
    procedure bFermeClick(Sender: TObject);
    // modif mbo 19.06.06 procedure ToolbarButton973Click(Sender: TObject);
    procedure BoutonAideClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormCreate(Sender: TObject);
  private
    { Déclarations privées }
    fCodeImmo : string;
    fDesImmo : string;
    fFiscal : boolean;
    fCurDateCalcul : string;
    TypeTraitement : integer;

    procedure InitZones;
    procedure RecalculDotations;
  public
    { Déclarations publiques }
  end;

//procedure AfficheDotationsIntermediaires (PlanAmor: TPlanAmort);
//procedure AfficheDotationsIntermediaires (CodeImmo : string);
// mbo 20.09.2006 ajout du paramètre pri pour affichage des dot prime
// mbo 18.10.2006 modif du paramètre en integer :
// valeur = 0 = plan d'amortissement
// valeur = 1 = suramortissement (non utilisé)
// valeur = 2 = plan d'amortissement subvention
procedure AfficheDotationsIntermediaires (CodeImmo,DesImmo : string; TypeAmort : integer);

implementation

uses Outils,ImOuPlan, ImPlanInfo;

{$R *.DFM}

//procedure AfficheDotationsIntermediaires (PlanAmor: TPlanAmort);
procedure AfficheDotationsIntermediaires (CodeImmo,DesImmo : string; TypeAmort : integer);
var  FDotInt: TFDotInt;

begin
  FDotInt := TFDotInt.Create (Application);
  FDotInt.fCodeImmo := CodeImmo;
  FDotInt.fDesImmo := DesImmo; //EPZ 03/11/00
  FDotInt.TypeTraitement := TypeAmort;
  try
    FDotInt.ShowModal;
  finally
    FDotInt.Free;
  end;
end;

procedure TFDotInt.FormShow(Sender: TObject);
begin
  fGrid.PostDrawCell := PostDrawCell;
  InitZones;
  RecalculDotations;
end;

procedure TFDotInt.InitZones;
begin
  fGrid.ColAligns[0]:= taCenter;
  fGrid.ColAligns[COL_CUMANT]:= taRightJustify;
  fGrid.ColAligns[COL_DOTCALC]:= taRightJustify;
  fGrid.ColAligns[COL_CUMAMORT]:= taRightJustify;
  fGrid.ColAligns[COL_VNC]:= taRightJustify;
  fDateCalcul.Text := DateToStr(VHImmo^.Encours.Fin);
end;

//=============================================================================

procedure TFDotInt.RecalculDotations;
var DateRef : TDateTime;
    PlanInfo : TPlanInfo;
    //CF,CE,AF,AE,EC, EA : double;
    DotDerog, ReintFisc, InfoMere, InfoFille : double;
    TvaReverser, PrixVente : double;

begin
  InfoMere := 0;
  InfoFille := 0;
  ReintFisc := 0;
  DateRef := StrToDate (fDateCalcul.Text);
  PlanInfo:= TPlanInfo.Create (fCodeImmo);
  PlanInfo.ChargeImmo(fCodeImmo);
  // BTY 11/05 FQ 16999 Cumul antérieur faux
  // mbo - 29/10/07 - fq 21754 - PlanInfo.Calcul(DateRef,True);
  PlanInfo.Calcul(DateRef,True, false, '');

if TypeTraitement = 1  then   // pour la prime = non utilisé
begin
  pAmortDerog.visible:=false;

  fGrid.Cells [COL_CUMANT,LIG_ECO] := StrfMontant ( PlanInfo.CumulAntPri,20, V_PGI.OkDecV , '', true);
  fGrid.Cells [COL_DOTCALC,LIG_ECO] := StrfMontant ( PlanInfo.DotationPri,20, V_PGI.OkDecV , '', true);
  fGrid.Cells [COL_CUMAMORT,LIG_ECO] := StrfMontant ( PlanInfo.CumulPri,20, V_PGI.OkDecV , '', true);
  fGrid.Cells [COL_VNC,LIG_ECO] := StrfMontant ( PlanInfo.VNCPri,20, V_PGI.OkDecV , '', true);

  fGrid.Cells [COL_CUMANT,LIG_FISC] := '';
  fGrid.Cells [COL_DOTCALC,LIG_FISC] := '';
  fGrid.Cells [COL_CUMAMORT,LIG_FISC] := '';
  fGrid.Cells [COL_VNC,LIG_FISC] := '';

end else if TypeTraitement = 0 then    // pour le comptable et le fiscal
begin
  fFiscal := PlanInfo.Plan.Fiscal;
  pAmortDerog.visible:=fFiscal;

  tReintDeduc.Visible := (PlanInfo.Plan.GestionFiscale) or (PlanInfo.Reintegration <> 0)
                         or (PlanInfo.QuotePart <> 0)
                         or (PlanInfo.Plan.EstRemplacee)
                         or (PlanInfo.Plan.Remplace);

                        // AND     (StrToDate (fDateCalcul.Text) = VHImmo^.Encours.Fin)) ; // ajout pour chantier fiscal

  ReintDeduc.Visible := tReintDeduc.Visible;
  TypeReint.Visible := tReintDeduc.Visible;

  InfoReint.Visible := PlanInfo.Plan.EstRemplacee;
  InfoDeduc.Visible := PlanInfo.Plan.Remplace;

  fGrid.Cells [0,1] := HM.Mess[0];
  fGrid.Cells [0,2] := HM.Mess[1];

  fGrid.Cells [COL_CUMANT,LIG_ECO] := StrfMontant ( PlanInfo.CumulAntEco,20, V_PGI.OkDecV , '', true);
  fGrid.Cells [COL_DOTCALC,LIG_ECO] := StrfMontant ( PlanInfo.DotationEco,20, V_PGI.OkDecV , '', true);
  fGrid.Cells [COL_CUMAMORT,LIG_ECO] := StrfMontant ( PlanInfo.CumulEco,20, V_PGI.OkDecV , '', true);
  fGrid.Cells [COL_VNC,LIG_ECO] := StrfMontant ( PlanInfo.VNCEco,20, V_PGI.OkDecV , '', true);

  //CE := PlanInfo.CumulEco;
  //AE := PlanInfo.CumulAntEco;

  if fFiscal then
  begin
    fGrid.Cells [COL_CUMANT,LIG_FISC] := StrfMontant ( PlanInfo.CumulAntFisc,20, V_PGI.OkDecV , '', true);
    fGrid.Cells [COL_DOTCALC,LIG_FISC] := StrfMontant ( PlanInfo.DotationFisc,20, V_PGI.OkDecV , '', true);
    fGrid.Cells [COL_CUMAMORT,LIG_FISC] := StrfMontant ( PlanInfo.CumulFisc,20, V_PGI.OkDecV , '', true);

    // ajout des calculs de derog par cumul et on différence entre dot fiscale et dot éco
    // FQ 15380 - mbo 31.03.2006
    { modif chantier fiscal fq 17512
    CF := PlanInfo.CumulFisc;
    AF := PlanInfo.CumulAntFisc;

    EC := Arrondi(CF-CE, V_PGI.OkDecV);
    EA := Arrondi(AF-AE,V_PGI.OkDecV);

    DotDerog := 0;
    if (EC >= 0) then
    begin
      if EA < 0 then EA := 0;
      if EC > EA then DotDerog := EC-EA
      else if EA > EC then DotDerog := EC-EA;
    end else
    if (EC < 0) then if (EA>0) then DotDerog := (-1)*EA;
    }

    // AE := CE; MVG 12/07/2006 Correction conseil, les vaiables ne servent plus aprés
    // AF := CF; MVG 12/07/2006 Correction conseil, les vaiables ne servent plus aprés

    { fq 15380 AmortDerog.Caption := StrfMontant (abs(PlanInfo.DotationFisc-PlanInfo.DotationEco),20, V_PGI.OkDecV , '', true);
    if (PlanInfo.DotationFisc-PlanInfo.DotationEco)>0 then AmortDerog.Caption := AmortDerog.Caption + ' '+HM.Mess[4]
    else if (PlanInfo.DotationFisc-PlanInfo.DotationEco)<0 then AmortDerog.Caption := AmortDerog.Caption + ' '+HM.Mess[5]; }

    InfoMere := 0;
    InfoFille := 0;

    if (PlanInfo.GetInfoCession((StrToDate (fDateCalcul.Text)),tvaReverser,PrixVente)= true) then
    begin
       DotDerog  := PlanInfo.GetCumulDerogatoireCede;
       ReintFisc := PlanInfo.GetCumulFECcede;
       ReintFisc := Arrondi((ReintFisc + PlanInfo.Reintegration + PlanInfo.QuotePart),V_PGI.OkDecV) ;

       ReintFisc := ReintFisc * (-1);

       If (PlanInfo.Plan.EstRemplacee) then
       begin
          infoMere := PlanInfo.Plan.GetValResiduelle(fCodeImmo,PlanInfo.Plan.AmortEco,StrToDate (fDateCalcul.Text)+1, true);
          ReintFisc:= ReintFisc + InfoMere;
       end;

       If (PlanInfo.Plan.Remplace) and (PlanInfo.Plan.DatedebEco <= DateRef) then
       begin
          InfoFille := PlanInfo.Plan.ValeurAchat;
          ReintFisc := ReintFisc - InfoFille;  //on déduit le montant HT
       end;

       DotDerog := DotDerog *(-1);

    end else
    begin
       DotDerog  := PlanInfo.GetDerogatoire;
       ReintFisc := PlanInfo.GetReintegration;
       ReintFisc := Arrondi((ReintFisc + PlanInfo.Reintegration + PlanInfo.QuotePart),V_PGI.OkDecV) ;

       if (PlanInfo.Plan.Remplace) and (PlanInfo.Plan.DatedebEco <= DateRef) then
       begin
          ReintFisc:= ReintFisc
             - PlanInfo.Plan.ValeurAchat;

          InfoFille := PlanInfo.Plan.ValeurAchat;
       end;
    end;

    AmortDerog.Caption := StrfMontant (abs(DotDerog),20, V_PGI.OkDecV , '', true);
    if DotDerog >0 then typederog.Caption := TraduireMemoire(' : Dotation')     //AmortDerog.Caption + ' '+HM.Mess[4]
    else if DotDerog <0 then  typederog.caption := TraduireMemoire(' : Reprise') // AmortDerog.Caption := AmortDerog.Caption + ' '+HM.Mess[5];
    else typederog.caption := '';

  end else
  begin

    fGrid.Cells [COL_CUMANT,LIG_FISC] := '';
    fGrid.Cells [COL_DOTCALC,LIG_FISC] := '';
    fGrid.Cells [COL_CUMAMORT,LIG_FISC] := '';
    fGrid.Cells [COL_VNC,LIG_FISC] := '';

    ReintFisc := Arrondi((ReintFisc + PlanInfo.Reintegration + PlanInfo.QuotePart),V_PGI.OkDecV) ;
  end;

   ReintDeduc.Caption := StrfMontant (abs(ReintFisc),20, V_PGI.OkDecV , '', true);
   if ReintFisc >0 then typereint.caption := TraduireMemoire(' : Réintégration') //ReintDeduc.Caption := ReintDeduc.Caption + ' '+HM.Mess[6]
   else if ReintFisc <0 then typereint.caption := TraduireMemoire(' : Déduction') //ReintDeduc.Caption := ReintDeduc.Caption + ' '+HM.Mess[7];
   else typereint.caption := '';

   InfoDeduc.Visible := false;
   InfoReint.Visible := false;

   if InfoMere<>0 then
   begin
      InfoReint.Visible := true;
      InfoReint.Caption      := TraduireMemoire('(dont réintégration de la VRC = ') +
                                 StrfMontant (abs(InfoMere),15, V_PGI.OkDecV , '', true)+')';
   end;

   if InfoFille<>0 then
   begin
      InfoDeduc.Visible := true;
      InfoDeduc.caption      := TraduireMemoire('(dont déduction du Montant HT = ') +
                                StrfMontant (abs(InfoFille),15, V_PGI.OkDecV , '', true)+')';
   end;

end else if TypeTraitement = 2 then    // plan d'amortissement subvention
begin
   // particularité pour gestion des antérieurs cumulés avec la première dotation
  pAmortDerog.visible:=false;

  // on cache les intitulés de la première colonne
  fGrid.Cells[0,0] := '';

  fGrid.Cells [COL_CUMANT,LIG_ECO] := StrfMontant ( PlanInfo.CumulAntSBV,20, V_PGI.OkDecV , '', true);
  fGrid.Cells [COL_DOTCALC,LIG_ECO] := StrfMontant ( PlanInfo.DotationSBV,20, V_PGI.OkDecV , '', true);
  fGrid.Cells [COL_CUMAMORT,LIG_ECO] := StrfMontant ( PlanInfo.CumulSBV,20, V_PGI.OkDecV , '', true);
  fGrid.Cells [COL_VNC,LIG_ECO] := StrfMontant ( PlanInfo.VNCsbv,20, V_PGI.OkDecV , '', true);

  fGrid.ColWidths[0] := -1;   // on masque la première colonne
  fGrid.rowcount := 2;        // on masque la deuxième ligne

end;

 PlanInfo.Free;
end;

procedure TFDotInt.PostDrawCell (Acol,ARow : LongInt ; Canvas : TCanvas;State : TGridDrawState) ;
var R : TRect ;
begin
  if (not fFiscal) and (ARow = 2) and (ACol > 0) then
  begin
    Canvas.Brush.Color := clBlack ;
    Canvas.Brush.Style := bsBDiagonal ;
    Canvas.Pen.Color   := clBlack ;
    Canvas.Pen.Mode    := pmCopy ;
    Canvas.Pen.Style   := psClear ;
    Canvas.Pen.Width   := 1 ;
    R:=fGrid.CellRect(ACol,ARow) ;
    Canvas.Rectangle(R.Left,R.Top,R.Right+1,R.Bottom+1) ;
  end;
end;

procedure TFDotInt.BImprimerClick(Sender: TObject);
begin
  {$IFDEF EAGLCLIENT}
   PrintDBGrid (Caption+' : '+fDesImmo,fGrid.Name,'');
  {$ELSE}
  PrintDBGrid (fGrid,Nil, Caption+' : '+fDesImmo,'');
  {$ENDIF}
end;

procedure TFDotInt.fDateCalculKeyPress(Sender: TObject; var Key: Char);
begin
  ParamDate (Self, Sender, Key);
  if key = #13 then FocusControl(fGrid);
end;

procedure TFDotInt.FormClose(Sender: TObject; var Action: TCloseAction);
begin fGrid.VidePile (False);end;

procedure TFDotInt.BValiderClick(Sender: TObject);
begin
  if not IsValidDate (fDateCalcul.text) then
  begin
    HM.Execute (3,Caption,'');
    FocusControl (fDateCalcul)
  end
  else if fDateCalcul.text <> fCurDateCalcul then
  begin
    if (StrToDate(fDateCalcul.Text) < VHImmo^.Encours.Deb) or
       ((StrToDate(fDateCalcul.Text) > VHImmo^.Suivant.Fin) and (VHImmo^.Suivant.Fin>iDate1900)) then
    begin
      HM.Execute (2,Caption,'');
      fDateCalcul.Text := DateToStr (VHImmo^.Encours.Fin);
      FocusControl (fDateCalcul);
    end
    else
    begin
    fCurDateCalcul:=fDateCalcul.text ;
    RecalculDotations;
    end ;
  end;
  ModalResult := mrNone;
end;

procedure TFDotInt.bFermeClick(Sender: TObject);
begin
  ModalResult := mrYes;
end;

// modif mbo 19.06.06 - fq 15848
procedure TFDotInt.BoutonAideClick(Sender: TObject);
begin
CallHelpTopic(Self);
end;

procedure TFDotInt.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
  case Key of
    VK_F9 :
      begin
        BValiderClick(nil);
        exit;
      end;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 02/06/2006
Modifié le ... :   /  /
Description .. : ajout de l'aide - fq 15848
Mots clefs ... :
*****************************************************************}
procedure TFDotInt.FormCreate(Sender: TObject);
begin
  HelpContext:=999999961 ;
end;



end.
