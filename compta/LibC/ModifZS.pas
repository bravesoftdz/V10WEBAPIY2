{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 25/01/2005
Modifié le ... :   /  /    
Description .. : Passage en eAGL
Mots clefs ... : 
*****************************************************************}
unit ModifZS;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, Ent1,
  StdCtrls, Hctrls, Buttons, ExtCtrls,
{$IFDEF EAGLCLIENT}
  uTob ,
{$ELSE}
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
  Grids, HEnt1, hmsgbox,HCompte,
  HSysMenu, HTB97,
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  ParamSoc //GetParamSoc {FP 29/12/2005}
  ,UentCommun  
  ;

function ModifZoneSerie(Table,Faxe : string; var Rafale:boolean): string;

type
  TFModifZS = class(TForm)
    Panel1: TPanel;
    BAjouter: TToolbarButton97;
    BEnlever: TToolbarButton97;
    FZModifiable: TGroupBox;
    FListe: TStringGrid;
    FModif: TGroupBox;
    FBox: TScrollBox;
    Panels_On: TPanel;
    MsgBox: THMsgBox;
    HMTrad: THSystemMenu;
    Dock: TDock97;
    HPB: TToolWindow97;
    BValider: TToolbarButton97;
    BAnnuler: TToolbarButton97;
    BAide: TToolbarButton97;
    BOK: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure BAjouterClick(Sender: TObject);
    procedure BEnleverClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure Panels_OnEnter(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure BOKClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private    { Déclarations privées }
    FTable,SQLUpdate,Faxe : string;
    CurLig            : LongInt;
    OkRafale          : boolean;
    Lefb : TFichierBase ;
    Chantier : Boolean ;
    Fermeture : boolean; //fb 17/11/2005 FQ14934
    procedure ChargeListe ;
    function  ConstruitUpdate : string;
    function  AffecteDataType(NomChamp : string) : String;
    procedure NewPosition;
    Function  TraiteCorrespBudLibre(St : String) : Boolean ;
    Function  AffecteZoomTable(St : String) : TZoomTable;
    Function  TraiteChantier(St : String) : Boolean ;
    Function  AffectettdeChoixCod(NomChamp : string) : String;
    Function  AffectettCleRepart : String ;
    function HasDoublon(const Name : String) : boolean;
    procedure Modifie(var SQL : String);
    function DeplaceItem(Search: String; Action : boolean): boolean;    {fb 17/11/2005 FQ14934}
    function OkCreateTHEdit(NomChamp: string): Boolean;                 {FP 29/12/2005}
    procedure AffecteTablette(NomChamp : string; Ed: THEdit);           {FP 29/12/2005}
    procedure ComponentExit(Sender: TObject);                           {FP 29/12/2005}
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

Uses
  {$IFDEF MODENT1}
  CPProcMetier,
  CPObjetGen,
  {$ENDIF MODENT1}
  UtilPGI;

function ModifZoneSerie(Table,Faxe : string; var Rafale:boolean): string;
var FModifZS : TFModifZS;
BEGIN
if _Blocage(['nrCloture'],False,'nrAucun') then Exit ;
Result:='';
FModifZS:=TFModifZS.Create(Application);
try
  FModifZS.FTable:=Table;
  FModifZS.Faxe:=Faxe ;
  if FModifZS.ShowModal=mrok then
     BEGIN
     Result:=FModifZS.SQLUpdate;
     Rafale:=FModifZS.OkRafale;
     END;
finally
  FModifZS.Free;
end;
END;

procedure TFModifZS.FormShow(Sender: TObject);
begin
  { FQ 18518 BVE 11.04.07 }
  Fermeture := true;
  { END FQ 18518 }
  PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
  SQLUpDate:=''; Panels_On.Visible:=false;
  if FTable='GENERAUX' then BEGIN Caption:= MsgBox.Mess[0] ; Lefb:=fbGene ; END else
 if FTable='TIERS' then BEGIN Caption:= MsgBox.Mess[1] ; Lefb:=fbAux ; END else
  if FTable='BUDGENE' then BEGIN Caption:= MsgBox.Mess[2] ; Lefb:=fbBudgen ; END else
  if FTable='BUDSECT' then
     BEGIN
     Caption:= MsgBox.Mess[5] ;
      Case FAxe[2] of
           '1' : Lefb:=fbBudSec1 ;
           '2' : Lefb:=fbBudSec2 ;
           '3' : Lefb:=fbBudSec3 ;
           '4' : Lefb:=fbBudSec4 ;
           '5' : Lefb:=fbBudSec5 ;
         End ;
     END else
  if FTable='BUDJAL' then BEGIN Caption:= MsgBox.Mess[6] ; Lefb:=fbBudJal ; END else
   if FTable='SECTION' then
      BEGIN
      Caption:= MsgBox.Mess[3] ;
      Case FAxe[2] of
           '1' : Lefb:=fbAxe1 ;
           '2' : Lefb:=fbAxe2 ;
           '3' : Lefb:=fbAxe3 ;
           '4' : Lefb:=fbAxe4 ;
           '5' : Lefb:=fbAxe5 ;
         End ;
      END else
  if FTable='SECTION' then
  begin
    Chantier:=VH^.Cpta[AxeToFb(Faxe)].Chantier ;
  end
  else
  if FTable = 'ECRITURE' then
  begin
    Lefb    := fbNone; 
    Caption := MsgBox.Mess[7];
  end;

{ FQ 15717 - CA - 26/04/2005 - Mode "Fiche par fiche" inaccessible si modification analytique => pose problème de mise à jour
  de la table analytique depuis la fiche des généraux }
  BValider.Visible :=  (not ((FTable = 'GENERAUX') and (V_PGI.ExtendedFieldSelection='1'))) and (not (FTable = 'ECRITURE'));

  ChargeListe;
  UpdateCaption(Self) ;
end;

Function TFModifZS.TraiteChantier(St : String) : Boolean ;
BEGIN
Result:=True ;
if(St='S_CHANTIER')Or(St='S_MAITREOEUVRE') then
   if Not Chantier then Result:=False ;
END ;

{b fb 17/11/2005 FQ14934}
function TFModifZS.DeplaceItem(Search : String; Action : boolean): boolean;
var
  i : Integer;
  P : TPanel;
  C : TControl;
  Lig : integer;
begin
  Result := False;

  for i := 0 to FBox.ControlCount-1 do begin
    if FBox.Controls[i] is TPanel then begin
      P := TPanel(FBox.Controls[i]);
      C := P.FindChildControl(Search);
      if C<>nil then begin
        Result := True;
        if Action then
          Exit
        else
          break;
      end;
    end;
  end;
  for Lig:= 0 to FListe.RowCount-1 do begin
    if Action then begin
      if FListe.Cells[0, Lig]=Search then begin
        FListe.Row:=Lig;
        BAjouterClick(nil);
        end;
     end
    else begin
      if (FListe.Cells[0, Lig]=Search) and (FListe.RowHeights[Lig]=0) then begin
          FListe.Row:=Lig;
          CurLig:=Lig;
          BEnleverClick(nil);
        end;
      end;
    end;
end;
{e fb 17/11/2005 FQ14934}
procedure TFModifZS.ChargeListe ;
{b FP 29/12/2005}
  procedure RemplirGrille(QC: TQuery; R  : integer);
  Var St : String ;
  begin
    FListe.ColWidths[0]:=0; FListe.ColWidths[1]:=0;
    {$IFDEF CCS3}
    //on a qu'un seul axe, donc on ne passe qu'une fois
    While (Not QC.EOF) and (R=0) do
    {$ELSE}
    While (Not QC.EOF) do
    {$ENDIF}
       BEGIN
       if FTable='SECTION' then
          BEGIN
          if Not TraiteChantier(QC.FindField('DH_NOMCHAMP').AsString)then
             BEGIN QC.Next ; Continue ; END ;
          END ;
       (*if Pos('_TABLE',QC.FindField('DH_NOMCHAMP').AsString)>0 then
          BEGIN
          if Not V_PGI.VersionDev then BEGIN QC.Next ; Continue ; END ;
          if Pos('BS_TABLE9',QC.FindField('DH_NOMCHAMP').AsString)>0 then BEGIN QC.Next ; Continue ; END ;
          END ;*)
       FListe.RowCount:=R+1; FListe.Row:=R;
       FListe.Cells[0,R]:=QC.FindField('DH_NOMCHAMP').AsString;
       FListe.Cells[1,R]:=QC.FindField('DH_TYPECHAMP').AsString;
       St:=QC.FindField('DH_LIBELLE').AsString ;
       if St='' then St:=QC.FindField('DH_NOMCHAMP').AsString ;
       FListe.Cells[2,R]:=St;
       Inc(R) ; QC.Next ;
       END ;
  end;
{e FP 29/12/2005}
Var St,fct : String ;
    QC : TQuery ;
    R  : integer;
    CompteAcc1, CompteAcc2 : string;
    i  : integer; 
BEGIN
if FTable='' then Exit ;
If V_PGI.ExtendedFieldSelection<>'' then Fct:=V_PGI.ExtendedFieldSelection else fct:='Z' ;
V_PGI.ExtendedFieldSelection:='' ;
//SG6 28.02.05 Gestion mode croisaxe
//RR 28/04/05 Pas suffisant ! Rajout de Fct='1' FQ 13824
if ( (VH^.AnaCroisaxe) and (FTable = 'GENERAUX') and (Fct='1')) then
begin
  FListe.ColWidths[0] := 0;
  FListe.ColWidths[1] := 0;
  FListe.RowCount := 1;
  FListe.Row := 0;
  FListe.Cells[0,0] := 'G_VENTILABLE';
  FListe.Cells[1,0] := 'BOOLEAN';
  FListe.Cells[2,0] := 'Ventilable (O/N)';
end
else
begin

{b fb 17/11/2005 FQ14934}
{if fTable = 'TIERS' then st := '(DH_PREFIXE="'+TableToPrefixe(fTable)+'" OR DH_PREFIXE="YTC")'
                    else st := 'DH_PREFIXE="'+TableToPrefixe(fTable)+'"';

QC:=OpenSQL('Select DH_NOMCHAMP, DH_TYPECHAMP, DH_LIBELLE, DH_CONTROLE, DH_PREFIXE '+
            ' From DECHAMPS Where CompteAcc1 '+st+' and DH_CONTROLE like "%'+fct+'%"',TRUE) ;}

CompteAcc1 := '';
CompteAcc2 := '';
if fTable = 'TIERS' then begin
  CompteAcc1 := '( ';
{ FQ 21455 BVE 06.11.07
  st := '( (DH_PREFIXE="'+TableToPrefixe(fTable)+'" OR DH_PREFIXE="YTC")'; } 
  st := '( (DH_PREFIXE="'+TableToPrefixe(fTable)+'")'; 
{ END FQ 21455 }
  CompteAcc2 := ' ) OR (DH_NOMCHAMP="YTC_SCHEMAGEN") ) ORDER BY DH_PREFIXE, DH_NUMCHAMP ' ;
 end
else
  st := 'DH_PREFIXE="'+TableToPrefixe(fTable)+'"';

QC:=OpenSQL('Select DH_NOMCHAMP, DH_TYPECHAMP, DH_LIBELLE, DH_CONTROLE, DH_PREFIXE '+
            ' From DECHAMPS Where '+CompteAcc1 + st+
            ' and DH_CONTROLE like "%'+fct+'%"' + CompteAcc2,TRUE) ;
{e fb 17/11/2005 FQ14934}
//St:=ExtractFields(FTable,'Z') ; QC:=OpenSQL(St,true);
RemplirGrille(QC, 0);    {b FP 29/12/2005}
Ferme(QC) ;

{b FP 29/12/2005 Modèle de restriction analytique}
if FTable='GENERAUX' then
  begin
  if VH^.AnaCroisaxe then      {Ajoute le modèle du premier axe croisé}
    QC:=OpenSQL('Select DH_NOMCHAMP, DH_TYPECHAMP, DH_LIBELLE, DH_CONTROLE, DH_PREFIXE'+
                ' From DECHAMPS'+
                ' Where DH_PREFIXE="'+TableToPrefixe('CLIENGENEMODELA')+'"'+
                '   And DH_NOMCHAMP = "CLA_CODE1"',TRUE)
  else
    QC:=OpenSQL('Select DH_NOMCHAMP, DH_TYPECHAMP, DH_LIBELLE, DH_CONTROLE, DH_PREFIXE'+
                ' From DECHAMPS'+
                ' Where DH_PREFIXE="'+TableToPrefixe('CLIENGENEMODELA')+'"'+
                '   And DH_NOMCHAMP like "CLA_CODE%"',TRUE);
  RemplirGrille(QC, FListe.RowCount);
  Ferme(QC);
  if VH^.AnaCroisaxe then
    begin
    R := FListe.RowCount-1;
    for i := 1 to MaxAxe do
      begin
      if GetParamSocSecur('SO_VENTILA' + IntToStr(i), False) then
        begin
        FListe.Cells[0,R] := 'CLA_CODE'+IntToStr(i);
        FListe.Cells[2,R] := TraduireMemoire('Modèle de restrictions axes croisés');  {FP 19/04/2006 FQ17739}
        break;
        end;
      end;
    end;
  end;
{e FP 29/12/2005}
end;

FListe.Row:=0; FListe.SetFocus ;
END ;

Function TFModifZS.TraiteCorrespBudLibre(St : String) : Boolean ;
BEGIN
Result:=(Pos('_CORRESP',St)>0) Or (Pos('G_BUDGENE',St)>0) Or (Pos('_TABLE',St)>0)
        Or (Pos('BJ_GENEATTENTE',St)>0) Or (Pos('BJ_SECTATTENTE',St)>0);
END ;

procedure TFModifZS.BAjouterClick(Sender: TObject);
var P  : TPanel;
    L  : TLabel;
    CC : TCheckBox;
    C  : THValComboBox;
    E  : TEdit;
    HC : THCritMaskEdit;
    FL : THNumEdit;
    ZT : THCpteEdit ;
    Ed : THEdit;              {FP 29/12/2005}
    Name, TC : String;
    OkCase,OkCombo,OkEdit, OKDate, OkFloat,OkZoomTab : boolean;
    i                     : integer;
    SauveLefb : TFichierBase ; //fb 17/11/2005 FQ14934
begin
  if CurLig<0 then Exit ;
  CurLig:=FListe.Row ;
  TC:=FListe.Cells[1,CurLig];
  if TC='' then Exit ;
  // 14699
  Name := FListe.Cells[0,CurLig];
  if (Name='YTC_INDEMNITE') or (Name='YTC_AVANTAGE') then begin
    if HasDoublon(Name) then begin
      PGIInfo('Une rémunération ne peut pas appartenir à ces deux zones en même temps ("Avantages en nature" et "Indemnités et remboursements"). ', Caption);
      exit;
    end;
  end;

  {b FP 21/02/2006}
  if (FTable = 'TIERS') and ((Name = 'T_CORRESP1') or (Name = 'T_CORRESP2'))
     and TCompensation.IsCompensation then
    begin
    if Name = TCompensation.GetChampPlan then
      begin
      PGIInfo('La modification en série n''est pas autorisée pour le plan utilisé pour la compensation.');
      exit;
      end;
    end;
  {e FP 21/02/2006}
  // Message d'avertissement pour les comptes généraux sur modif G_IAS14 // Modif IFRS
  if ( FTable = 'GENERAUX' ) and (FListe.Cells[0,CurLig] = 'G_IAS14') then
    PGIInfo('Attention ! L''indicateur de traitement IAS14 ne sera pas mis à jour pour les comptes de nature : "Immobilisation", "Charge" ou "Produit".', Caption) ;
  // FIN Modif IFRS
  OkZoomTab:=TraiteCorrespBudLibre(FListe.Cells[0,CurLig]) ;
  { Ajout ME le 19/04/2001 }
  // ajout me est sur un car, il faut saisir comme un boolean
  if {(ctxPCL in V_PGI.PGIContexte) and }
  ((FListe.Cells[0,CurLig] = 'G_VENTILABLE1') or
   (FListe.Cells[0,CurLig] = 'G_VENTILABLE2') or
   (FListe.Cells[0,CurLig] = 'G_VENTILABLE3') or
   (FListe.Cells[0,CurLig] = 'G_VENTILABLE4') or
   (FListe.Cells[0,CurLig] = 'G_VENTILABLE5'))
  then TC :='BOOLEAN';
  { FQ 18568 BVE 11.04.07 }
  if (FListe.Cells[0,CurLig] = 'G_CONSO') then TC := 'VARCHAR(3)'; 
  { END FQ 18568 }
  OkCase:=(TC='BOOLEAN');
  OkCombo:=(TC='COMBO');
  OKDate:=(TC='DATE');
  OkFloat:=(TC='DOUBLE');
  OkEdit:=((TC<>'BOOLEAN') And (TC<>'COMBO') And (TC<>'DATE') And (TC<>'DOUBLE') And (Not OkZoomTab));
  if ((Not OkCase) And (Not OkCombo) And (Not OkEdit) And (Not OKDate) And (Not OkFloat) And (Not OkZoomTab)) then Exit;
{b fb 17/11/2005 FQ14934}
  if FListe.Cells[0,CurLig]='YTC_SCHEMAGEN' then begin
    OkEdit:=false;
    OkZoomtab:=true;
    end;
{e fb 17/11/2005 FQ14934}
  P:=TPanel.Create(Self); P.Parent:=FBox; P.Name:='P'+IntToStr(CurLig);
  P.ParentFont:=False; P.Font.Color:=clWindowText;
  P.Height:=30; P.Align:=alTop; P.BevelInner:=bvLowered; P.BevelOuter:=bvRaised;
  P.Caption:=''; P.OnEnter:=Panels_OnEnter; P.OnClick:=Panels_OnEnter;

  if OkCase then BEGIN
    CC:=TCheckBox.Create(Self); CC.Parent:=P; CC.Name:=FListe.Cells[0,CurLig];
    CC.Left:=8; CC.Width:=300; CC.Alignment:= taLeftJustify; CC.Top:=4;
    CC.Caption:=FListe.Cells[2,CurLig]; CC.setfocus;
    END
  else BEGIN
    L:=THLabel.Create(Self); L.Parent:=P; L.Name:='L'+IntToStr(CurLig);
    L.Left:=8; L.Width:=150; L.Top:=8; L.AutoSize:=false;
    L.Caption:=FListe.Cells[2,CurLig];
    // FQ 12399
    if OkDate then begin
      HC:=THCritMaskEdit.Create(Self); HC.Parent:=P; HC.Name:=FListe.Cells[0,CurLig];
      HC.Left:=L.Left+L.Width+1; HC.Width:=P.Width-HC.Left-8; HC.Top:=4;
      HC.Text := '';
      HC.EditMask:='!99/99/0000;1;_'; HC.ControlerDate := True;
      L.FocusControl:=HC; HC.SetFocus;
      end else
    if OKFloat then begin
      FL:=THNumEdit.Create(Self); FL.Parent:=P; FL.Name:=FListe.Cells[0,CurLig];
      FL.Left:=L.Left+L.Width+1; FL.Width:=P.Width-FL.Left-8; FL.Top:=4;
      L.FocusControl:=FL; FL.SetFocus;
      end else
    if OkCombo then BEGIN
      C:=THValComboBox.Create(Self); C.Parent:=P; C.Name:=FListe.Cells[0,CurLig];
      C.Left:=L.Left+L.Width+1; C.Width:=P.Width-C.Left-8; C.Top:=4; C.Text:='' ;
      C.Style:=csDropDownList ; C.DataType:=AffecteDataType(C.Name);
      {JP 30/10/07 : FQ 19150 : ajout du aucun, comme dans CPTIERS_TOM.PAS}
      if (UpperCase(C.Name) = 'T_RELANCETRAITE') then begin
        C.Vide := True;
        C.VideString := TraduireMemoire('Aucune');
        C.ReLoad;
      end;
      C.ItemIndex:=0; L.FocusControl:=C; C.SetFocus;
      END else
      {b FP 29/12/2005: Pour utiliser la property DataType au lieu de ZoomTable}
    if OkCreateTHEdit(FListe.Cells[0,CurLig]) then BEGIN
      Ed:=THEdit.Create(Self); Ed.Parent:=P; Ed.Name:=FListe.Cells[0,CurLig];
      Ed.Left:=L.Left+L.Width+1; Ed.Width:=P.Width-Ed.Left-8; Ed.Top:=4;
      Ed.Text:='';L.FocusControl:=Ed; Ed.SetFocus;
      Ed.OnExit := ComponentExit;
      AffecteTablette(FListe.Cells[0,CurLig], Ed);
      END else
      {e FP 29/12/2005}
    if OkEdit then BEGIN
      E:=TEdit.Create(Self); E.Parent:=P; E.Name:=FListe.Cells[0,CurLig];
      E.Left:=L.Left+L.Width+1; E.Width:=P.Width-E.Left-8; E.Top:=4;
      E.Text:=''; L.FocusControl:=E; E.SetFocus;
      END else
    if OkZoomTab then BEGIN
      {b fb 17/11/2005 FQ 14934}
      SauveLefb := Lefb;
      if FListe.Cells[0,CurLig]='YTC_SCHEMAGEN' then
        Lefb := fbGene;
      {e fb 17/11/2005 FQ 14934}
      ZT:=THCpteEdit.Create(Self); ZT.Parent:=P; ZT.Name:=FListe.Cells[0,CurLig];
      ZT.Left:=L.Left+L.Width+1; ZT.Width:=P.Width-ZT.Left-8; ZT.Top:=4;
      ZT.Text:=''; ZT.ZoomTable:=AffecteZoomTable(ZT.Name) ; L.FocusControl:=ZT; ZT.SetFocus;
      Lefb := SauveLefb;  //fb 17/11/2005 FQ 14934
    END;
  END;
  FListe.RowHeights[CurLig]:=0;
  NewPosition;
  for i:=0 to FBox.ControlCount-1 do
    if FBox.Controls[i] is TPanel then begin
      if TPanel(FBox.Controls[i])=P then
        TPanel(FBox.Controls[i]).Color:=clTeal
      else
        TPanel(FBox.Controls[i]).Color:=clBtnFace;
      end;
{b fb 17/11/2005 FQ14934}
  if (Name='YTC_SCHEMAGEN') then
    DeplaceItem('YTC_ACCELERATEUR', true);
  if (Name='YTC_ACCELERATEUR') then
    DeplaceItem('YTC_SCHEMAGEN', true);
{e fb 17/11/2005 FQ14934}
end;

procedure TFModifZS.BEnleverClick(Sender: TObject);
var P : TPanel;
    T : TTabOrder;
    i : integer;
begin
  P:= TPanel(FindComponent('P'+IntToStr(CurLig)));
  if P=Nil then exit;
  T:=P.TabOrder;
  while P.ControlCount>0 do
    P.Controls[0].Free;
  P.Free;
  FListe.RowHeights[CurLig]:=FListe.DefaultRowHeight;
  for i:=0 to FBox.ControlCount-1 do
    if FBox.Controls[i] is TPanel then begin
      if TPanel(FBox.Controls[i]).TabOrder=T-1 then
        TPanel(FBox.Controls[i]).Color:=clTeal
      else
        TPanel(FBox.Controls[i]).Color:=clBtnFace;
      end;
{b fb 17/11/2005 FQ14934}
  if (FListe.Cells[0,CurLig]='YTC_SCHEMAGEN') then
    DeplaceItem('YTC_ACCELERATEUR', false);
  if (FListe.Cells[0,CurLig]='YTC_ACCELERATEUR') then
    DeplaceItem('YTC_SCHEMAGEN', false);
{e fb 17/11/2005 FQ14934}
end;

procedure TFModifZS.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
var P : TPanel;
begin
{b fb 17/11/2005 FQ14934}
  if not Fermeture then
    CanClose:=false;
  if CanClose then begin
{e fb 17/11/2005 FQ14934}
    While FBox.ControlCount>0 do
       BEGIN
       P:= TPanel(FBox.Controls[0]);
       while P.ControlCount>0 do P.Controls[0].Free ;
       P.Free;
       END;
    end;  //fb 17/11/2005 FQ14934
end;

procedure TFModifZS.Panels_OnEnter(Sender: TObject);
var St     : string;
    i,IP,j : integer;
    P      : TPanel;
begin
St:=TComponent(Sender).Name; if Pos('P',St)<=0 then exit;
CurLig:=StrToInt(Copy(St,2,Length(St)));
for i:=0 to FBox.ControlCount-1 do
   if ((FBox.Controls[i] is TPanel) and (FBox.Controls[i].Name<>'Panels_On')) then
      BEGIN
      P:=TPanel(FBox.Controls[i]); St:=P.Name; if Pos('P',St)<=0 then break;
      IP:=StrToInt(Copy(St,2,Length(St)));
      if IP=CurLig then
         BEGIN
         P.Color:=clTeal;
         for j:=0 to P.ControlCount-1 do
            BEGIN
            if P.Controls[j] is TCheckBox then BEGIN TCheckBox(P.Controls[j]).setfocus; break; END else
            if P.Controls[j] is TEdit then BEGIN TEdit(P.Controls[j]).setfocus; break; END else
            if P.Controls[j] is THEdit then BEGIN THEdit(P.Controls[j]).setfocus; break; END else {FP 29/12/2005} 
            if P.Controls[j] is THValComboBox then BEGIN THValComboBox(P.Controls[j]).setfocus; break; END else
            if P.Controls[j] is THCpteEdit then BEGIN THCpteEdit(P.Controls[j]).setfocus; break; END else
            if P.Controls[j] is THCritMaskEdit then BEGIN THCritMaskEdit(P.Controls[j]).setfocus; break; END else
            if P.Controls[j] is THNumEdit then BEGIN THNumEdit(P.Controls[j]).setfocus; break; END;
            END;
         END else P.Color:=clBtnFace;
      END;
end;

procedure TFModifZS.FListeDblClick(Sender: TObject);
begin BAjouterClick(NIL); end;

function TFModifZS.ConstruitUpdate : string;
var i,j : integer;
    P   : TPanel;
    Champ,Valeur : String;
    C            : TControl;
{b fb 17/11/2005 FQ14934}
    MesCaseAccCouple   : boolean;
    CompteCurseurAcc   : integer;
    ValeurCompteAcc    : string;
    LibAcc             : string;
    OkAcc              : boolean;
{e fb 17/11/2005 FQ14934}
BEGIN
  Result:='';
 {b fb 17/11/2005 FQ14934}
  MesCaseAccCouple:=false;
  CompteCurseurAcc:=-1;
 {e fb 17/11/2005 FQ14934}
  For i:=0 to FBox.ControlCount-1 do BEGIN
    if (FBox.Controls[i] is TPanel) and (FBox.Controls[i].Name<>'Panels_On') then BEGIN
      P:=TPanel(FBox.Controls[i]);
      For j:=0 to P.ControlCount-1 do BEGIN
        C:=P.Controls[j];
        if C is TLabel then {RAF}
        else BEGIN
          Champ:=C.Name;
          if C is TCheckBox then BEGIN
            if TCheckBox(C).Checked then
              if Pos('CONFIDENTIEL',Champ)<>0 then Valeur:='1' else Valeur:='X'
            else
              if Pos('CONFIDENTIEL',Champ)<>0 then Valeur:='0' else Valeur:='-' ;
{b fb 17/11/2005 FQ14934}
            if (TCheckBox(C).Checked) and (Champ='YTC_ACCELERATEUR') then
              MesCaseAccCouple:=true;
{e fb 17/11/2005 FQ14934}
            END else
          if C is THValComboBox then BEGIN
            Valeur:=THValComboBox(C).Value;
            END else
          if C is TEdit then BEGIN
            Valeur:=TEdit(C).Text;
{b fb 17/11/2005 FQ14934}
            if (Champ='YTC_SCHEMAGEN') then begin
              CompteCurseurAcc  := i;
              ValeurCompteAcc   := Valeur;

              if not MesCaseAccCouple then begin
                Valeur :=BourreOuTronque(Valeur, fbGene);
                if ValeurCompteAcc<>'' then begin
                  LibAcc := RechDom('TZGENERAL', ValeurCompteAcc, False, 'G_FERME <> "X"');
                  if (LibAcc = '') then
                    Valeur:='';
                  end;
                end;
              end;
{e fb 17/11/2005 FQ14934}
            END else
{b FP 29/12/2005}
          if C is THEdit then BEGIN
            Valeur:=THEdit(C).Text;
            END else
{e FP 29/12/2005}
          if C is THNumEdit then BEGIN
            Valeur:=THNumEdit(C).Text;
            END else
          if C is THCritMaskEdit then
          BEGIN
            // GCO - 21/11/2005 - FQ 17039
            if IsValidDate( THCritMaskEdit(C).Text ) then
              Valeur := UsDateTime(StrToDate(THCritMaskEdit(C).Text))
            else
              Valeur := THCritMaskEdit(C).Text;
          END else
          if C is THCpteEdit then BEGIN
            Valeur:=THCpteEdit(C).Text;
          END ;
        END;
      END;
      if i<FBox.ControlCount-1 then Result:=Result+Champ+'="'+Valeur+'",;'
                               else Result:=Result+Champ+'="'+Valeur+'" ';
    END;
  END;
{b fb 17/11/2005 FQ14934}
  Fermeture:=true;
  OkAcc    := true;
  if MesCaseAccCouple and OkRafale then begin
    ValeurCompteAcc :=BourreOuTronque(ValeurCompteAcc, fbGene);

    if ValeurCompteAcc='' then begin
      MsgBox.Mess.Add('8;Comptes auxiliaires;Le compte d''accélération que vous avez renseigné n'#39'est pas un compte général.;W;O;O;O;');
      MsgBox.execute(8,'','');
      OkAcc:=false;
      end
    else begin
      LibAcc := RechDom('TZGENERAL', ValeurCompteAcc, False, 'G_FERME <> "X"');
      if (LibAcc = '') then begin
        MessageAlerte('Le compte : '+ ValeurCompteAcc + ' est fermé');
        OkAcc:=false;
        end
      else
        OkAcc:=true;
      end;

    for i:=0 to FBox.ControlCount-1 do
      if FBox.Controls[i] is TPanel then begin
        if i=CompteCurseurAcc then
          TPanel(FBox.Controls[i]).Color:=clTeal
        else
          TPanel(FBox.Controls[i]).Color:=clBtnFace;
        end;
    end;
  Fermeture:=OkAcc;
{e fb 17/11/2005 FQ14934}

  // GCO - 31/10/2006 - FQ 18942
  if Fermeture then
  begin
    if Champ = 'T_PAYEUR' then
    begin
      if ExisteSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE = "' + Valeur + '"') then
      begin
        if GetColonneSQL('TIERS', 'T_ISPAYEUR', 'T_AUXILIAIRE = "' + Valeur + '"') <> 'X' then
        begin
          PgiInfo('Le compte auxiliaire ' + Valeur +  ' n''est pas déclaré comme compte tiers payeur/payé.', Caption);
          Fermeture := False;
        end;
      end
      else
      begin
        PgiInfo('Le compte auxiliaire ' + Valeur + ' n''existe pas.', Caption);
        Fermeture := False;
      end;
    end;
  end;
END;

procedure TFModifZS.BValiderClick(Sender: TObject);
begin
FListe.SetFocus;             {FP 29/12/2005 Force l'évènement OnExit}
if TToolBarButton97(Sender).Name='BValider' then OkRafale:=False else OkRafale:=True ;
SQLUpdate:=ConstruitUpdate;
if (FTable = 'TIERS') then Modifie(SQLUpdate); // 14713
end;

procedure TFModifZS.BOKClick(Sender: TObject);
begin BValiderClick(Sender) ; end;

function TFModifZS.AffecteDataType(NomChamp : string) : String;
BEGIN
Result:='';
if FTable='SECTION' then
   if(NomChamp='S_CLEREPARTITION') then
      BEGIN Result:=AffectettCleRepart ; Exit ; END ;
If FTable='TIERS' Then If NomChamp='T_MODEREGLE' Then BEGIN Result:='TTMODEREGLE' ; Exit ;END ;
Result:=Get_Join(NomChamp);
if Result='' then Result:=AffectettdeChoixCod(NomChamp) ;
END;

Function TFModifZS.AffectettCleRepart : String;
BEGIN
Case Lefb of
     fbAxe1 : Result:='ttCleRepart1' ;
     fbAxe2 : Result:='ttCleRepart2' ;
     fbAxe3 : Result:='ttCleRepart3' ;
     fbAxe4 : Result:='ttCleRepart4' ;
     fbAxe5 : Result:='ttCleRepart5' ;
     else Result:='' ;
   End ;
END ;

Function TFModifZS.AffectettdeChoixCod(NomChamp : string) : String;
BEGIN
Result:='';
if NomChamp='G_TPF' then BEGIN Result:='ttTPF' ; Exit ; END ;
if(NomChamp='G_QUALIFQTE1') Or (NomChamp='G_QUALIFQTE2') then
   BEGIN Result:='ttQualUnitMesure' ; Exit ; END ;
if(NomChamp='G_RELANCEREGLEMENT') Or (NomChamp='T_RELANCEREGLEMENT') then
   BEGIN Result:='ttRelanceRegle' ; Exit ; END ;
if(NomChamp='G_RELANCETRAITE') Or (NomChamp='T_RELANCETRAITE') then
   BEGIN Result:='ttRelanceTraite' ; Exit ; END ;
if(NomChamp='G_MOTIFVIREMENT') Or (NomChamp='T_MOTIFVIREMENT') then
   BEGIN Result:='ttNatureEco' ; Exit ; END ;
if(NomChamp='G_LETTREPAIEMENT') Or (NomChamp='T_LETTREPAIEMENT') then
   BEGIN Result:='ttLettrePaiement' ; Exit ; END ;
if(NomChamp='G_REGIMETVA') Or (NomChamp='T_REGIMETVA') then
   BEGIN Result:='ttRegimeTva' ; Exit ; END ;
if(NomChamp='T_TVAENCAISSEMENT') then BEGIN Result:='ttTVAEncaissement' ; Exit ; END ;
END ;

Function TFModifZS.AffecteZoomTable(St : string) : TZoomTable ;
Var i : Byte ;
BEGIN
Result:=tzGeneral ;
if Pos('_CORRESP',St)>0 then
   BEGIN
   i:=StrToInt(Copy(St,Length(St),1)) ;
   Case Lefb of
        fbGene    : if i=1 then Result:=tzCorrespGene1 else Result:=tzCorrespGene2 ;
        fbAux     : if i=1 then Result:=tzCorrespAuxi1 else Result:=tzCorrespAuxi2 ;
        fbAxe1    : if i=1 then Result:=tzCorrespSec11 else Result:=tzCorrespSec12 ;
        fbAxe2    : if i=1 then Result:=tzCorrespSec21 else Result:=tzCorrespSec22 ;
        fbAxe3    : if i=1 then Result:=tzCorrespSec31 else Result:=tzCorrespSec32 ;
        fbAxe4    : if i=1 then Result:=tzCorrespSec41 else Result:=tzCorrespSec42 ;
        fbAxe5    : if i=1 then Result:=tzCorrespSec51 else Result:=tzCorrespSec52 ;
       End ;
   END ;
if Pos('_TABLE',St)>0 then
   BEGIN
   if St[1]='G' then
      BEGIN
      Case St[Length(St)] of
           '0' : Result:=tzNatGene0 ; '1' : Result:=tzNatGene1 ; '2' : Result:=tzNatGene2 ;
           '3' : Result:=tzNatGene3 ; '4' : Result:=tzNatGene4 ; '5' : Result:=tzNatGene5 ;
           '6' : Result:=tzNatGene6 ; '7' : Result:=tzNatGene7 ; '8' : Result:=tzNatGene8 ;
           '9' : Result:=tzNatGene9 ;
          End ;
      END else
      if St[1]='S' then
         BEGIN
         Case St[Length(St)] of
              '0' : Result:=tzNatSect0 ; '1' : Result:=tzNatSect1 ; '2' : Result:=tzNatSect2 ;
              '3' : Result:=tzNatSect3 ; '4' : Result:=tzNatSect4 ; '5' : Result:=tzNatSect5 ;
              '6' : Result:=tzNatSect6 ; '7' : Result:=tzNatSect7 ; '8' : Result:=tzNatSect8 ;
              '9' : Result:=tzNatSect9 ;
             End ;
         END else
         if St[1]='T' then
            BEGIN
            Case St[Length(St)] of
                 '0' : Result:=tzNatTiers0 ; '1' : Result:=tzNatTiers1 ; '2' : Result:=tzNatTiers2 ;
                 '3' : Result:=tzNatTiers3 ; '4' : Result:=tzNatTiers4 ; '5' : Result:=tzNatTiers5 ;
                 '6' : Result:=tzNatTiers6 ; '7' : Result:=tzNatTiers7 ; '8' : Result:=tzNatTiers8 ;
                 '9' : Result:=tzNatTiers9 ;
                End ;
            END else
            if St[2]='G' then
               BEGIN
               Case St[Length(St)] of
                    '0' : Result:=tzNatBud0 ; '1' : Result:=tzNatBud1 ; '2' : Result:=tzNatBud2 ;
                    '3' : Result:=tzNatBud3 ; '4' : Result:=tzNatBud4 ; '5' : Result:=tzNatBud5 ;
                    '6' : Result:=tzNatBud6 ; '7' : Result:=tzNatBud7 ; '8' : Result:=tzNatBud8 ;
                    '9' : Result:=tzNatBud9 ;
                   End ;
               END else
               if St[2]='S' then
                  BEGIN
               Case St[Length(St)] of
                    '0' : Result:=tzNatBudS0 ; '1' : Result:=tzNatBudS1 ; '2' : Result:=tzNatBudS2 ;
                    '3' : Result:=tzNatBudS3 ; '4' : Result:=tzNatBudS4 ; '5' : Result:=tzNatBudS5 ;
                    '6' : Result:=tzNatBudS6 ; '7' : Result:=tzNatBudS7 ; '8' : Result:=tzNatBudS8 ;
                    '9' : Result:=tzNatBudS9 ;
                   End ;
                  END ;
   END ;
if St='BJ_GENEATTENTE' then Result:=tzBudgenAtt ;
if St='BJ_SECTATTENTE' then
   Case Faxe[2] of
        '1' : Result:=tzBudSecAtt1 ;
        '2' : Result:=tzBudSecAtt2 ;
        '3' : Result:=tzBudSecAtt3 ;
        '4' : Result:=tzBudSecAtt4 ;
        '5' : Result:=tzBudSecAtt5 ;
      End ;
if (Pos('_TABLE',St)<=0) And (Pos('_CORRESP',St)<=0) And
   (Pos('BJ_GENEATTENTE',St)<=0) And (Pos('BJ_SECTATTENTE',St)<=0) then
   BEGIN
   Case Lefb of
        fbGene    : Result:=tzGeneral ; fbAux     : Result:=tzTiers ;
        fbAxe1    : Result:=tzSection ; fbAxe2    : Result:=tzSection2 ;
        fbAxe3    : Result:=tzSection3 ; fbAxe4    : Result:=tzSection4 ;
        fbAxe5    : Result:=tzSection5 ; fbBudgen  : Result:=tzBudgen   ;
        fbBudSec1 : Result:=tzBudSec1  ; fbBudSec2 : Result:=tzBudSec2  ;
        fbBudSec3 : Result:=tzBudSec3  ; fbBudSec4 : Result:=tzBudSec4  ;
        fbBudSec5 : Result:=tzBudSec5  ; fbBudJal  : Result:=tzBudJal   ;
       End ;
   END ;
END ;

procedure TFModifZS.NewPosition;
var i, New : integer;
BEGIN
New:=CurLig;
for i:=CurLig to FListe.RowCount-2 do
   if FListe.RowHeights[i]=0 then New:=New+1 else break;
if FListe.RowHeights[New]=0 then
   BEGIN
   New:=CurLig;
   for i:=CurLig downto 1 do
      if FListe.RowHeights[i]=0 then New:=New-1 else break;
   END;
if FListe.RowHeights[New]=0 then CurLig:=-1 else BEGIN CurLig:=New; FListe.Row:=New; END;
END;

procedure TFModifZS.BAnnulerClick(Sender: TObject);
begin
  Fermeture:=true; //fb 17/11/2005 FQ14934
  Close ;
end;

procedure TFModifZS.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

// 14699
function TFModifZS.HasDoublon(const Name : String): boolean;
var
  szSearch : String;
  i : Integer;
  P : TPanel;
  C : TControl;
begin
  Result := False;

  if (Name='YTC_INDEMNITE') then szSearch := 'YTC_AVANTAGE' else
    if (Name='YTC_AVANTAGE') then szSearch := 'YTC_INDEMNITE' else exit;

  for i := 0 to FBox.ControlCount-1 do begin
    if FBox.Controls[i] is TPanel then begin
      P := TPanel(FBox.Controls[i]);
      C := P.FindChildControl(szSearch);
      if C <> nil then begin
        Result := True;
        exit;
      end;
    end;
  end;
end;

// 14713
procedure TFModifZS.Modifie(var SQL: String);
var
  bControlOtherValue, bDAS2, bProfession, bRemuneration, bIndemnite, bAvantage : Boolean;
  Value : String;

  // Retourne la valeur du champ passé en paramètre
  function GetValue(const Name : String) : String;
  var
    i : Integer;
  begin
    Result := '';
    i := Pos(Name, SQL);
    if i <=0 then exit
    else begin
      Result := Copy(SQL, i+Length(Name)+2, Length(SQL));
      i := Pos(',;', Result);
      Result := Copy(Result, 1, i-2);
    end;
  end;

  // Supprime la valeur du champ passé en paramètre
  procedure DelValue(const Name : String);
  var
    i : Integer;
    sz, sz2 : String;
  begin
    i := Pos(Name, SQL);
    if i<=0 then exit
    else begin
      sz := Copy(SQL, 1, i+Length(Name)+1); // Pour garder la séquence ="
      sz2 := Copy(SQL, Length(sz)+1, Length(SQL)-Length(sz));
      i := Pos('"', sz2);
      SQL := sz+Copy(sz2, i, Length(sz2)-i+1); // Pour garder la séquence ",;
    end;
  end;

begin
  bControlOtherValue := False;
  bDAS2         := (Pos('YTC_DAS2',         SQL)>0);
  bProfession   := (Pos('YTC_PROFESSION',   SQL)>0);
  bRemuneration := (Pos('YTC_REMUNERATION', SQL)>0);
  bIndemnite    := (Pos('YTC_INDEMNITE',    SQL)>0);
  bAvantage     := (Pos('YTC_AVANTAGE',     SQL)>0);

  if bDAS2 then begin
    Value := GetValue('YTC_DAS2');
    // Si DAS2 coché
    if Value = 'X' then bControlOtherValue := True
    // Si DAS2 décoché : Met à blanc les autres valeurs
    else begin
      if bProfession   then DelValue('YTC_PROFESSION')   else SQL := SQL+',;YTC_PROFESSION=""';
      if bRemuneration then DelValue('YTC_REMUNERATION') else SQL := SQL+',;YTC_REMUNERATION=""';
      if bIndemnite    then DelValue('YTC_INDEMNITE')    else SQL := SQL+',;YTC_INDEMNITE=""';
      if bAvantage     then DelValue('YTC_AVANTAGE')     else SQL := SQL+',;YTC_AVANTAGE=""';
    end;
    end
  else bControlOtherValue := True;


  if bControlOtherValue then begin
    Value := GetValue('YTC_REMUNERATION');
    if (Value='AVA') then begin // Avantage en nature
      if bIndemnite    then DelValue('YTC_INDEMNITE')    else SQL := SQL+',;YTC_INDEMNITE=""';
      end else
    if (Value='IND') then begin // Indemnités et remboursements
      if bAvantage     then DelValue('YTC_AVANTAGE')     else SQL := SQL+',;YTC_AVANTAGE=""';
      end
    else begin
      if bIndemnite    then DelValue('YTC_INDEMNITE')    else SQL := SQL+',;YTC_INDEMNITE=""';
      if bAvantage     then DelValue('YTC_AVANTAGE')     else SQL := SQL+',;YTC_AVANTAGE=""';
    end;
  end;
end;
{b FP 29/12/2005}
function TFModifZS.OkCreateTHEdit(NomChamp: string): Boolean;
begin
  Result := Pos('CLA_CODE', NomChamp) <> 0;       {FP 20/04/2006}
end;

procedure TFModifZS.AffecteTablette(NomChamp: string; Ed: THEdit);
begin
  if Pos('CLA_CODE', NomChamp) <> 0 then          {FP 20/04/2006}
    begin
    Ed.DataType := 'CMODELRESTANA';
    Ed.Plus     := 'CRA_AXE="A'+NomChamp[Length(NomChamp)]+'"';
    Ed.ElipsisButton   := True;
    Ed.ElipsisAutoHide := True;
    end;
end;

procedure TFModifZS.ComponentExit(Sender: TObject);
var
  Axe:      String;
  Code:     String;
  CompName: String;
  SQL:      String;
begin
  CompName := (Sender as TComponent).Name;
  if Pos('CLA_CODE', CompName) <> 0 then          {FP 20/04/2006}
    begin
    Axe  := 'A'+CompName[Length(CompName)];
    Code := Trim((Sender as THEdit).Text);
    SQL  := 'select CRA_CODE from CMODELRESTANA'+
            ' where CRA_AXE="'+Axe+'"'+
            '   and CRA_CODE="'+Code+'"';
    if (Code <> '') and (not ExisteSQL(SQL)) then
      begin
      (Sender as THEdit).ElipsisClick(Sender);
      (Sender as THEdit).SetFocus;
      ModalResult := mrNone;
      end;
    end;
end;
{e FP 29/12/2005}

end.
