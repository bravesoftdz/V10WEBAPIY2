// modif mbo - 02.01.2007 - FQ 19422 - si pas d'échéance suite import de crédit bail : erreur
//                                     ajout d'un test + message
unit ListeEch;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  hmsgbox, HTB97, Grids, Hctrls, ExtCtrls, HPanel, HEnt1, HSysMenu, UiUtil,
  StdCtrls,ImEnt;

procedure  ListeDesEcheances (ListeEcheance : TList; Residuel : double);

type
  TFListeEche = class(TForm)
    HPanel1: THPanel;
    HPanel2: THPanel;
    GridEcheancier: THGrid;
    Dock971: TDock97;
    PBouton: TToolWindow97;
    BFerme: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    HPanel3: THPanel;
    HM: THMsgBox;
    HMTrad: THSystemMenu;
    BImprimer: TToolbarButton97;
    HLabel1: THLabel;
    LoyerProg: THNumEdit;
    LoyerGen: THNumEdit;
    EngageGen: THNumEdit;
    HLabel3: THLabel;
    EngageProg: THNumEdit;
    procedure FormShow(Sender: TObject);
    procedure GridEcheancierRowEnter(Sender: TObject; Ou: Integer;
      var Cancel: Boolean; Chg: Boolean);
    procedure BImprimerClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure HelpBtnClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Déclarations privées }
    ListeEcheance : TList;
    Residuel : double;
    procedure AfficheCumulEcheancier(LigneSel : integer);
    procedure DessineCell(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
  public
    { Déclarations publiques }
  end;

implementation

uses  PlanEche,
      {$IFDEF EAGLCLIENT}
      utileAGL
      {$ELSE}
      PrintDBG
      {$ENDIF}
      ;

{$R *.DFM}

procedure  ListeDesEcheances (ListeEcheance : TList; Residuel : double);
var
  FListeEche: TFListeEche;
  PP : THPanel;
begin
  // ajout mbo pour correction fiche 19422 - 02.01.2007
  if ListeEcheance.count = 0 then
  begin
     PGIBox('Aucune échéance saisie pour ce bien', '');
  end else
  begin
    FListeEche := TFListeEche.Create(Application);
    FListeEche.ListeEcheance := ListeEcheance;
    FListeEche.Residuel := Residuel;
    PP:=FindInsidePanel ;
    if PP=Nil then
     BEGIN
      try
         FListeEche.ShowModal;
      finally
         FListeEche.Free ;
      end ;
     END else
     BEGIN
     InitInside(FListeEche,PP) ;
     FListeEche.Show ;
     END ;
  end;
end;

procedure TFListeEche.FormShow(Sender: TObject);
var
   i : integer;
   ARecord : PEcheance;
   Engagement : double;
begin
   inherited;
   // Initialisation des grids
   GridEcheancier.GetCellCanvas:=DessineCell;
   GridEcheancier.ColAligns[0]:= taRightJustify;
   GridEcheancier.ColAligns[1]:= taCenter;
   GridEcheancier.ColAligns[2]:= taRightJustify;
   GridEcheancier.ColAligns[3]:= taRightJustify;
   // Affichage de l'échéancier
   if ListeEcheance <> nil then
   begin
     GridEcheancier.RowCount := ListeEcheance.Count+1;
     Engagement := 0.0;
     for i:=0 to ListeEcheance.Count - 1 do
     begin
       ARecord := ListeEcheance.Items [i];
       Engagement := Engagement + ARecord^.Montant;
     end;
     LoyerGen.Value:= Engagement;
     EngageGen.Value:= Residuel;
     Engagement := Engagement + Residuel;
     for i:=0 to ListeEcheance.Count - 1 do
     begin
       ARecord := ListeEcheance.Items [i];
       GridEcheancier.Cells[0,i+1] := IntToStr (i+1);
       GridEcheancier.Cells[1,i+1] := DateToStr (ARecord^.Date);
       GridEcheancier.Cells[2,i+1] := strfMontant (ARecord^.Montant,20,V_PGI.OkDecV,'',false);
       Engagement := Engagement - ARecord^.Montant;
       GridEcheancier.Cells[3,i+1] := strfMontant (Engagement,20,V_PGI.OkDecV,'',false);
     end;
     AfficheCumulEcheancier(1);
   end
   else
   AfficheCumulEcheancier(0);

   // FQ 16296 tga 31/08/2005
   LoyerGen.masks.PositiveMask := StrfMask(V_PGI.OkDecV,'', True);
   EngageGen.masks.PositiveMask := StrfMask(V_PGI.OkDecV,'', True);
   // fin tga 31/08/2005

end;

procedure TFListeEche.AfficheCumulEcheancier(LigneSel : integer);
var i: integer;
    TotalLoyer : double;
begin
  TotalLoyer := 0.0;
  for i := 1 to LigneSel do
      TotalLoyer := TotalLoyer + StrToFloat(GridEcheancier.Cells [2,i]);
//  GridCumul.Cells[1,0]:= StrfMontant (TotalLoyer,20,V_PGI.OkDecV,'',false);
  LoyerProg.Value:= TotalLoyer;
//  if LigneSel <> 0 then GridCumul.Cells[2,0]:= GridEcheancier.Cells [3,LigneSel]
  if LigneSel <> 0 then EngageProg.Value := StrToFloat(GridEcheancier.Cells [3,LigneSel])
//  else GridCumul.Cells[2,0]:= StrfMontant (0.00,20,V_PGI.OkDecV,'',false);
  else EngageProg.Value:= 0.00;

  // FQ 16296 tga 31/08/2005
  LoyerProg.masks.PositiveMask := StrfMask(V_PGI.OkDecV,'', True);
  EngageProg.masks.PositiveMask := StrfMask(V_PGI.OkDecV,'', True);
  // fin tga 31/08/2005

end;

procedure TFListeEche.GridEcheancierRowEnter(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
  AfficheCumulEcheancier(Ou);
end;

procedure TFListeEche.BImprimerClick(Sender: TObject);
begin
  {$IFDEF EAGLCLIENT}
  PrintDBGrid (Caption,GridEcheancier.Name,'');  
  {$ELSE}
  PrintDBGrid (GridEcheancier,Nil, Caption,'');
  {$ENDIF}
end;

procedure TFListeEche.FormClose(Sender: TObject; var Action: TCloseAction);
begin
GridEcheancier.VidePile (False);
if isInside(Self) then Action:=caFree ;
end;

procedure TFListeEche.DessineCell(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
var DateEcheance : TDateTime;
    bIsInExercice : boolean;
begin
  if ARow = 0 then exit;
  DateEcheance := StrToDate(GridEcheancier.Cells[1,ARow]);
  bIsInExercice := (DateEcheance >= VHImmo^.Encours.Deb) and (DateEcheance <= VHImmo^.Encours.Fin);
  //bIsPaye := (DateEcheance < V_PGI.DateEntree);
  //if bIsPaye then Canvas.Font.Style := Canvas.Font.Style + [fsItalic];
  if bIsInExercice then Canvas.Font.Style := Canvas.Font.Style + [fsBold];
end ;

procedure TFListeEche.HelpBtnClick(Sender: TObject);
begin
CallHelpTopic(Self);
end;

procedure TFListeEche.FormCreate(Sender: TObject);
begin
{$IFDEF SERIE1}
HelpContext:=511020 ;
{$ELSE}
HelpContext:=2101300 ;
{$ENDIF}
end;

end.
