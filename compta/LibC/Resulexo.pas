{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 29/10/2004
Modifié le ... :   /  /
Description .. : - CA - 29/10/2004 - FQ 10944 - Réécriture du calcul des
Suite ........ : soldes
Mots clefs ... :
*****************************************************************}
unit Resulexo;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, Hctrls, StdCtrls, Buttons, ExtCtrls, HTB97, HPanel, Uiutil, HSysMenu,
{$IFDEF EAGLCLIENT}
  uTob, utileAGL, //GraphUtil,
{$ELSE}
  DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  PrintDBG, MajTable, Graph,
{$ENDIF}
  Hqry,
  HEnt1,
  Ent1,
  SAISUTIL,
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ELSE}
  tCalcCum,
  {$ENDIF MODENT1}
  CpteUtil ;

Procedure ResultatDeLexo ;

type
  TFResulexo = class(TForm)
    Fliste   : THGrid;
    HPB: TToolWindow97;
    Pcrit    : TPanel;
    LExo     : TLabel;
    LDevise  : TLabel;
    LType    : TLabel;
    TEtab    : THLabel;
    CbExo    : THValComboBox;
    CbDevise : THValComboBox;
    CbTypMvt : THValComboBox;
    CbEtab   : THValComboBox;
    Calculsolde: THNumEdit;
    CbMess: TComboBox;
    BGraph: TToolbarButton97;
    HMTrad: THSystemMenu;
    Rev: TCheckBox;
    FListe1: THGrid;
    BStop: TToolbarButton97;
    BImprimer: TToolbarButton97;
    BChercher: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    Dock: TDock97;
    Timer1: TTimer;
    CBFinDate: TComboBox;
    CBDebDat: TComboBox;
    procedure FormShow(Sender: TObject);
    procedure CbExoChange(Sender: TObject);
    procedure CbDeviseChange(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure CbEtabChange(Sender: TObject);
    procedure BGraphClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BStopClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure Timer1Timer(Sender: TObject);
  private
    ExerciceEntree : TExoDate ;
    NbrMois : Word ;
    Devise : RDevise ;
    FiltreDev,FiltreEtab,FiltreExo,OkStop : Boolean ;
    CodEtb : String ;
    WMinX,WMinY : Integer ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure ResultatCumuler ;
    Procedure AffichePeriode ;
{$IFDEF EAGLCLIENT}
    Function GenererTOBEtat : TOB ;
{$ELSE}
    Procedure ResultatCumuler_UDF ;
    procedure CreateViewResult ;
{$ENDIF}
  public
    {JP 28/06/06 : FQ 16149 : gestion des réstrictions Etablissements et à défaut des ParamSoc}
    procedure GereEtablissement;
  end;


implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPProcGen,
  CPProcMetier,
  ULibExercice,
  {$ENDIF MODENT1}
  HStatus;

Procedure ResultatDeLexo ;
var FResulexo : TFResulexo ;
    PP : THPanel ;
BEGIN
FResulexo:=TFResulexo.Create(Application) ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     FResulexo.ShowModal ;
    Finally
     FResulexo.Free ;
    End ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FResulexo,PP) ;
   FResulexo.Show ;
   END ;
END ;

procedure TFResulexo.FormShow(Sender: TObject);
begin
Calculsolde.Visible:=False ; Rev.Visible:=V_PGI.Controleur ;
if Rev.Visible then
   BEGIN Rev.Checked:=True ; Rev.State:=cbGrayed ; END else Rev.Checked:=False ;
CbExo.Value:=EXRF(VH^.Entree.Code) ;
CbTypMvt.Value:='NOR' ;
CbDevise.ItemIndex:=0 ;
CbEtab.ItemIndex:=0 ;
Devise.Code:='' ;

{JP 28/06/06 : FQ 16149 : refonte de la gestion des établissements
PositionneEtabUser(cbEtab,True) ;}
GereEtablissement;

GetInfosDevise(Devise) ;
{$IFDEF EAGLCLIENT}
BGraph.Visible := False ;
{$ELSE}
if ((V_PGI.UDF) and (V_PGI.Driver<>dbMSACCESS)) then CreateViewResult ;
{$ENDIF}
if estSerie(S5) then Rev.Visible:=False ;
end;

Procedure TFResulexo.AffichePeriode ;
Var i : Byte ;
    DebDate,FinDate : TDateTime ;
    StM : String ;
BEGIN
  DebDate:=ExerciceEntree.Deb ;
  CBDebDat.Items.Clear ;
  CBFinDate.Items.Clear ;
  CBDebDat.Items.Add('') ;
  CBFinDate.Items.Add('') ;

  for i:=1 to NbrMois do
  BEGIN
    // GCO - 08/09/2004 FQ 14337
{    if i = NbrMois then
      FinDate := ExerciceEntree.Fin
    else
      FinDate := FindeMois(DebDate);}

    FinDate := FindeMois(DebDate) ;
    CBDebDat.Items.Add(DatetoStr(DebDate)) ;
    CBFinDate.Items.Add(DatetoStr(FinDate)) ;
    StM := FormatDateTime('mmmm yyyy',DebDate) ;
    StM[1] := upcase(StM[1]) ;
    FListe.Cells[0,i] := StM ;
    Fliste.RowCount := Fliste.RowCount+1 ;
    DebDate := DebutDeMois(PlusMois(DebDate,1));
  END ;

  Fliste.Cells[0,Fliste.RowCount-1] := CbMess.Items[1] ;
  Fliste.Row :=1 ;
  FListe1.RowCount := FListe.RowCount ;
END ;

procedure TFResulexo.CbExoChange(Sender: TObject);
begin
ExerciceEntree.Code:=CbExo.Value ; RempliExoDate(ExerciceEntree) ;
NbrMois:=ExerciceEntree.NombrePeriode ;
Caption:=CbMess.Items[2]+' : '+CbExo.Text ;
FiltreExo:=(CbExo.Value<>'') ;
UpdateCaption(Self) ;
END ;

procedure TFResulexo.CbDeviseChange(Sender: TObject);
begin
if CbDevise.Value='' then Devise.Code:=''else Devise.Code:=CbDevise.Value ;
GetInfosDevise(Devise) ; FiltreDev:=(CbDevise.Value<>'') ;
end;

Procedure TFResulexo.ResultatCumuler ;
Var Q : TQuery ;
    i,j : Byte ;
    Natgen : String ;
    SDeb,SCre,Totaldeb,TotalCre : Extended ;
    DebDate,FinDate : TDateTime ;
    ToTal : TabTot ;
    NbTour : Integer ;
    SetTypePie : SetttTypePiece ;
BEGIN
Fliste.RowCount:=2 ; AffichePeriode ;
NbTour:=2*NbrMois ; InitMove(NbTour,CbMess.Items[0]) ;
SetTypePie:=WhatTypeEcr(CbTypMvt.Value,V_PGI.Controleur,Rev.State) ;
for i:=1 to 2 do    //Parceque Charge et Produit
   BEGIN
   if i=1 then NatGen:='CHA' else NatGen:='PRO' ;
   Q:=PrepareTotCptJointure(fbGene,SetTypePie,FiltreDev,FiltreEtab,FiltreExo,FALSE,'G_','GENERAL','G_NATUREGENE',NatGen);
   Totaldeb:=0 ; TotalCre:=0 ;
   for j:=1 to NbrMois do
      BEGIN
      Application.ProcessMessages ;
      if OkStop then Break ;
      DebDate:=StrToDate(CBDebDat.Items[j]) ;
      FinDate:=StrToDate(CBFinDate.Items[j]) ;
      ExecuteTotCptJointure(Q,DebDate,FinDate,Devise.Code,CodEtb,ExerciceEntree.Code,Total,True,V_PGI.OkDecV,V_PGI.OkDecE,
                            fbGene,SetTypePie,FALSE,'G_','GENERAL','G_NATUREGENE',NatGen) ;
//      SDeb:=0 ; SCre:=0 ;
      SDeb:=Total[1].TotDebit ; SCre:=Total[1].TotCredit ;
      Totaldeb:=Totaldeb+SDeb ; TotalCre:=TotalCre+SCre ;
      AfficheLeSolde(Calculsolde,SDeb,SCre) ;
      FListe.Cells[i,j]:=Calculsolde.Text ;
      MoveCur(False) ;
      END ;
   AfficheLeSolde(Calculsolde,Totaldeb,TotalCre) ;
   FListe.Cells[i,Fliste.RowCount-1]:=Calculsolde.Text ;
   If Q<>NIL Then Ferme(Q) ;
   if OkStop then Break ;
   END ;
//Totaldeb:=0 ; TotalCre:=0 ;
{ CA - 29/10/2004 - Réécriture du calcul des soldes pour FQ 10944 }
(*
for i:=1 to FListe.RowCount-1 do
  BEGIN
  SDeb:=Valeur(Fliste.Cells[1,i]) ;
  SCre:=Valeur(Fliste.Cells[2,i]) ;
  if (Pos('D',Fliste.Cells[1,i])>0) And (Pos('D',Fliste.Cells[2,i])>0) then
     BEGIN SDeb:=SDeb+SCre ; SCre:=0 ; END ;
  if (Pos('C',Fliste.Cells[1,i])>0) And (Pos('C',Fliste.Cells[2,i])>0) then
     BEGIN SCre:=SCre+SDeb ; SDeb:=0 ; END ;
  if SCre<=0 then
   AfficheLeSolde(Calculsolde,SDeb,-SCre)else
    BEGIN STemp:=SDeb ; SDeb:=Scre ; SCre:=Stemp ; AfficheLeSolde(Calculsolde,SDeb,SCre) ; END ;
  FListe.Cells[3,i]:=Calculsolde.Text ;
  END ;
  *)
  for i:=1 to FListe.RowCount-1 do
  begin
    SDeb:=Valeur(Fliste.Cells[1,i]) ;
    SCre:=Valeur(Fliste.Cells[2,i]) ;
    if ((SCre <= 0) and (SDeb >= 0)) then AfficheLeSolde(Calculsolde,SDeb,(-1)*SCre)
    else if ((SCre <= 0) and (SDeb <= 0)) then AfficheLeSolde(Calculsolde,0,((-1)*SCre)+((-1)*SDeb))
    else if ((SCre >= 0) and (SDeb <= 0)) then AfficheLeSolde(Calculsolde,SCre+SDeb,0)
    else if ((SCre >= 0) and (SDeb >= 0)) then AfficheLeSolde(Calculsolde,SDeb+SCre,0);
    FListe.Cells[3,i]:=Calculsolde.Text ;
  end;
FiniMove ; Fliste.SetFocus ;
END ;

{$IFDEF EAGLCLIENT}

Function TFResulexo.GenererTOBEtat : TOB ;
var iRow  : Integer ;
    lTob  : TOB ;
begin
  Result := TOB.Create('_TOBRESULTAT', nil, -1) ;
  For iRow := 1 to Fliste.RowCount - 1 do
    begin
    lTob := TOB.Create('TOBRESULTAT', Result, -1) ;
    lTob.AddChampSupValeur('REX_INDEX',    iRow );                 // index pour tri
    lTob.AddChampSupValeur('REX_PERIODE',  FListe.Cells[0,iRow] ); // période
    lTob.AddChampSupValeur('REX_CHARGE',   FListe.Cells[1,iRow] ); // total charges
    lTob.AddChampSupValeur('REX_PRODUIT',  FListe.Cells[2,iRow] ); // total produits
    lTob.AddChampSupValeur('REX_RESULTAT', FListe.Cells[3,iRow] ); // total résultats
    end ;
end ;

{$ELSE}
procedure TFResulexo.CreateViewResult ;
Var St : String ;
BEGIN
if TableExiste('MOISECR') then
   BEGIN
   if _ChampExiste('MOISECR','E_MOIS') then Exit ;
    Try
     BeginTrans ;
     ExecuteSQL('DROP VIEW MOISECR') ;
     CommitTrans ;
    Except
     RollBack ;
    end ;
   END ;
St:='CREATE VIEW MOISECR '+
    '(E_MOIS,E_DEBIT,E_CREDIT,E_DEBITDEV,E_CREDITDEV,G_NATUREGENE,E_DATECOMPTABLE,E_EXERCICE,E_ECRANOUVEAU,E_QUALIFPIECE,E_DEVISE,E_ETABLISSEMENT) ' +
    'AS SELECT '+
    DB_Month('E_DATECOMPTABLE')+',E_DEBIT,E_CREDIT,E_DEBITDEV,E_CREDITDEV,G_NATUREGENE,E_DATECOMPTABLE,E_EXERCICE,E_ECRANOUVEAU,E_QUALIFPIECE,E_DEVISE,E_ETABLISSEMENT '+
    'FROM ECRITURE,GENERAUX WHERE G_GENERAl=E_GENERAL' ;
 Try
  BeginTrans ;
  ExecuteSQL(St) ;
  CommitTrans ;
 Except
  RollBack ;
 End ;
END ;

Procedure TFResulexo.ResultatCumuler_UDF ;
Var Q : TQuery ;
    i,j : Byte ;
    StSDev,StQul,StEtb,StDev : String ;
    SDeb,SCre,STemp,TotaldebC,TotalCreC,TotaldebP,TotalCreP : Extended ;
    //DebDate,FinDate : TDateTime ;
BEGIN
Fliste.RowCount:=2 ; AffichePeriode ;
AfficheLeSolde(Calculsolde,0,0) ;
for i:=1 to NbrMois+1 do BEGIN FListe.Cells[1,i]:=Calculsolde.Text ; FListe.Cells[2,i]:=Calculsolde.Text ; FListe.Cells[3,i]:=Calculsolde.Text ;END ;
if FiltreDev then StSDev:='SUM(E_DEBITDEV),SUM(E_CREDITDEV)' else StSDev:='SUM(E_DEBIT),SUM(E_CREDIT)' ;
if FiltreDev then StDev:='AND E_DEVISE="'+Devise.Code+'" ' else StDev:='' ;
if FiltreEtab then StDev:='AND E_ETABLISSEMENT="'+CodEtb+'" ' else StEtb:='' ;
StQul:=WhereSupp('E_',WhatTypeEcr(CbTypMvt.Value,FALSE,Rev.State)) ;
Q:=OpenSQL('SELECT G_NATUREGENE,E_MOIS,'+StSDev+' '+
           'From MOISECR '+
           'WHERE E_EXERCICE="'+ExerciceEntree.Code+'" '+
           StDev+StEtb+StQul+
           'AND (G_NATUREGENE="CHA" or G_NATUREGENE="PRO") '+
           'Group by E_MOIS,G_NATUREGENE',TRUE) ;
InitMove(48,CbMess.Items[0]) ;
TotaldebC:=0 ; TotalCreC:=0 ; TotaldebP:=0 ; TotalCreP:=0 ;
While Not Q.EOF do
   BEGIN
   j:=Q.Fields[1].AsInteger ; if Q.Fields[0].AsString='CHA' then i:=1 else i:=2 ;
   SDeb:=Q.Fields[2].AsFloat ; ; SCre:=Q.Fields[3].AsFloat ; ;
   //DebDate:=StrToDate(CBDebDat.Items[j]) ;
   //FinDate:=StrToDate(CBFinDate.Items[j]) ;
   MoveCur(False) ;
   if i=1 then BEGIN TotaldebC:=TotaldebC+SDeb ; TotalCreC:=TotalCreC+SCre ; END
          else BEGIN TotaldebP:=TotaldebP+SDeb ; TotalCreP:=TotalCreP+SCre ; END ;
   AfficheLeSolde(Calculsolde,SDeb,SCre) ;
   FListe.Cells[i,j]:=Calculsolde.Text ;
   Q.Next ;
   END ;
AfficheLeSolde(Calculsolde,TotaldebC,TotalCreC) ; FListe.Cells[1,Fliste.RowCount-1]:=Calculsolde.Text ;
AfficheLeSolde(Calculsolde,TotaldebP,TotalCreP) ; FListe.Cells[2,Fliste.RowCount-1]:=Calculsolde.Text ;
Ferme(Q) ;
//TotaldebC:=0 ; TotalCreC:=0 ;
for i:=1 to FListe.RowCount-1 do
  BEGIN
  SDeb:=Valeur(Fliste.Cells[1,i]) ;
  SCre:=Valeur(Fliste.Cells[2,i]) ;
  if (Pos('D',Fliste.Cells[1,i])>0) And (Pos('D',Fliste.Cells[2,i])>0) then
     BEGIN SDeb:=SDeb+SCre ; SCre:=0 ; END ;
  if (Pos('C',Fliste.Cells[1,i])>0) And (Pos('C',Fliste.Cells[2,i])>0) then
     BEGIN SCre:=SCre+SDeb ; SDeb:=0 ; END ;
  if SCre<=0 then
   AfficheLeSolde(Calculsolde,SDeb,-SCre)else
    BEGIN STemp:=SDeb ; SDeb:=Scre ; SCre:=Stemp ; AfficheLeSolde(Calculsolde,SDeb,SCre) ; END ;
  FListe.Cells[3,i]:=Calculsolde.Text ;
  END ;
FiniMove ; Fliste.SetFocus ;
END ;
{$ENDIF}

procedure TFResulexo.BChercherClick(Sender: TObject);
begin
OkStop:=False ;
if Not V_PGI.OutLook then BEGIN PCrit.Enabled:=False ; BGraph.Enabled:=False ; END ;
ChangeMask(Calculsolde,Devise.Decimale,'') ;
{$IFNDEF EAGLCLIENT}
if ((V_PGI.UDF) and (V_PGI.Driver<>dbMSACCESS)) then
  ResultatCumuler_UDF
else
{$ENDIF}
  ResultatCumuler ;
if Not V_PGI.OutLook then BEGIN PCrit.Enabled:=True ; BGraph.Enabled:=True ; END ;
// Possibilité d'imprimer après avoir effectué les calculs
BImprimer.Enabled := true ;
end;

procedure TFResulexo.CbEtabChange(Sender: TObject);
begin CodEtb:=CbEtab.Value ; FiltreEtab:=(CodEtb<>'') ; end;

procedure TFResulexo.BGraphClick(Sender: TObject);
{$IFDEF EAGLCLIENT}
{$ELSE}
Var i,j : Integer ;
    St  : String ;
{$ENDIF}
begin
{$IFDEF EAGLCLIENT}
{$ELSE}
{ Liste cachée avec resultat inversé pour afficher dans le grapheur }
  For i:=1 to Fliste.RowCount-2 do
    For j:=1 to Fliste.ColCount-1 do
      BEGIN
        St:=Fliste.Cells[j,i] ;
        if Pos('D',St)>0 then St[Pos('D',St)]:='C' else
        if Pos('C',St)>0 then St[Pos('C',St)]:='D' ;
        FListe1.Cells[j,i]:=St ;
      END ;
VisuGraph('RESULEXO',Caption,FListe1,CBDebDat.Items,0,1,FListe1.RowCount-2,TRUE,[3,1,2],NIL,Nil) ;
{$ENDIF}
end;

procedure TFResulexo.BFermeClick(Sender: TObject);
var lBoDansPanel : Boolean ;
begin
  lBoDansPanel := IsInside(Self) ;
  Close ;
  if lBoDansPanel
    then CloseInsidePanel(Self) ;
end;

procedure TFResulexo.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFResulexo.FormCreate(Sender: TObject);
begin WMinX:=Width ; WMinY:=389 ; PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ; end;

procedure TFResulexo.FormResize(Sender: TObject);
begin
if V_PGI.OutLook then Exit ;
Fliste.ColWidths[1]:=(Fliste.Width - 126) div 3  ;
Fliste.ColWidths[2]:=(Fliste.Width - 126) div 3 ;
Fliste.ColWidths[3]:=(Fliste.Width - 126) div 3 ;
end;

procedure TFResulexo.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFResulexo.BStopClick(Sender: TObject);
begin
OkStop:=True ;
end;

procedure TFResulexo.BImprimerClick(Sender: TObject);
{$IFDEF EAGLCLIENT}
var lTOBEtat : TOB ;
{$ENDIF}
begin
{$IFDEF EAGLCLIENT}
  lTOBEtat := GenererTOBEtat ;
	LanceEtatTOB( 'E' , 'REX' , 'REX' ,
                lTOBEtat,
                True, False, False, Nil, '', Caption, False ) ;
  lTOBEtat.Free ;
{$ELSE}
  PrintDBGrid(FListe,Nil,Caption,'') ;
{$ENDIF}
end;

procedure TFResulexo.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if IsInside(Self) then Action:=caFree ;
end;

procedure TFResulexo.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  // F10 ou F9
  if (Key=VK_F10) or (Key=VK_F9) then
    begin
    Key:=0 ;
    BChercherClick(Nil) ;
    end ;
end;

procedure TFResulexo.Timer1Timer(Sender: TObject);
begin
Timer1.Enabled:=False ;
if ctxPCL in V_PGI.PGIContexte then BChercherClick(Nil) ;
end;

{---------------------------------------------------------------------------------------}
procedure TFResulexo.GereEtablissement;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(cbEtab) then begin
    {Si l'on ne gère pas les établissement ...}
    if not VH^.EtablisCpta  then begin
      {... on affiche l'établissement par défaut}
      cbEtab.Value := VH^.EtablisDefaut;
      {... on désactive la zone}
      cbEtab.Enabled := False;
    end

    {On gère l'établisement, donc ...}
    else begin
      {... On commence par regarder les restrictions utilisateur}
      PositionneEtabUser(cbEtab);
      {... s'il n'y a pas de restrictions, on reprend le paramSoc
       JP 25/10/07 : FQ 19970 : Finalement on oublie l'option de l'établissement par défaut
      if cbEtab.Value = '' then begin
        {... on affiche l'établissement par défaut
        cbEtab.Value := VH^.EtablisDefaut;
        {... on active la zone
        cbEtab.Enabled := True;
      end;}
    end;
  end;

end;

end.
