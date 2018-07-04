{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 17/10/2017
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTPARAMLIGBAST ()
Mots clefs ... : TOF;BTPARAMLIGBAST
*****************************************************************}
Unit BTPARAMLIGBAST_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     HSysMenu,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,FE_main,
{$else}
     eMul, 
{$ENDIF}
     LookUp,
     Vierge,
     uTob, 
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Hpanel,
     UTOF,
     HTB97 ;

Type
  TOF_BTPARAMLIGBAST = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    TOBCat : TOB;
    GS : THGrid;
    ListeChamps :string;
    HMTrad: THSystemMenu;

    //
    procedure CATLIGNEClick (Sender :TObject);
    procedure ChargeGrille;
    procedure ActiveEventsGrid(Status : boolean);
    procedure ActiveEventSaisie(Status : boolean);
    procedure PrepareAffichage;
    procedure ChargeTOBviaCat(Categorie : string);
    procedure AfficheGrid;
    function GetTOBCat (TOBCat : TOB; Ligne : integer) : TOB;
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure PlusMoinsClick (Sender : TObject);
    procedure SearchRefErp (Sender : TObject);
    procedure ChangeType (Sender : TObject);
    procedure BeforeChange(Ou: Integer);
    procedure PARAMBSVClick (Sender : TObject);


  end ;

Implementation
uses AGLInitGC;

procedure TOF_BTPARAMLIGBAST.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMLIGBAST.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMLIGBAST.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMLIGBAST.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMLIGBAST.OnArgument (S : String ) ;
begin
  Inherited ;
  TOBCat := TOB.Create('LES DETAIL CATEGORIE',nil,-1);
  GS := THGrid(GetControl('GS'));
  HmTrad := TFVierge(Ecran).HMTrad;
  //
  THValComboBox (GetControl('BM6_CATLIGNE')).OnChange := CATLIGNEClick;
  THValComboBox (GetControl('BM6_CATLIGNE')).Value := '001';
  TToolbarButton97 (GetControl('BPARAMBSV')).onclick := PARAMBSVClick;
end ;

procedure TOF_BTPARAMLIGBAST.OnClose ;
begin
  BeforeChange (GS.Row);
  TOBCat.free;
  Inherited ;
end ;

procedure TOF_BTPARAMLIGBAST.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMLIGBAST.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTPARAMLIGBAST.CATLIGNEClick(Sender: TObject);
begin
  BeforeChange (GS.Row) ;
  ChargeGrille;
end;

procedure TOF_BTPARAMLIGBAST.BeforeChange (Ou : Integer);
var TOBL : TOB;
    P1 : THPanel;
begin
  if Ou < 1 then Exit;
  if Pos(THValComboBox (GetControl('BM6_TYPEERP')).Value,'CAU;RG')>0 then
  begin
    THEdit(GetControl('BM6_ARTICLE')).Text := '';
  end;
  TOBL := GetTOBCat(TOBCat,Ou); if TOBL = nil then Exit;
  P1 := THpanel(GetControl('P1'));
  TOBL.GetEcran(Ecran,P1);
  if TOBL.IsOneModifie(false) then
  begin
    TOBL.UpdateDB;
    TOBL.SetAllModifie(false);
  end;
end;

procedure TOF_BTPARAMLIGBAST.ChargeGrille;
var cancel : Boolean;
begin
  cancel := false;
  ActiveEventsGrid(false);
  ActiveEventSaisie(false);
  GS.VidePile(false);
  GS.RowCount := 1;
  ChargeTOBviaCat(THValComboBox (GetControl('BM6_CATLIGNE')).Value);
  PrepareAffichage;
  AfficheGrid;
  if Gs.rowCount = 1 then GS.rowcount := 2;
  GS.Row := 1;
  GSRowEnter(Self,GS.row,cancel,false);
  ActiveEventsGrid(true);
end;

procedure TOF_BTPARAMLIGBAST.ActiveEventsGrid(Status: boolean);
begin
  if Status then
  begin
    GS.OnRowEnter := GSRowEnter;
    GS.OnRowExit := GSRowExit;
  end else
  begin
    GS.OnRowEnter := nil;
    GS.OnRowExit := nil;
  end;
end;

procedure TOF_BTPARAMLIGBAST.ChargeTOBviaCat(Categorie: string);
begin
  TOBCat.ClearDetail;
  TOBCat.LoadDetailDBFromSQL('BTYPELIGBAST','SELECT *,"" AS MODIFIED FROM BTYPELIGBAST WHERE BM6_CATLIGNE="'+Categorie+'" ORDER BY BM6_INDICE',false);
end;

procedure TOF_BTPARAMLIGBAST.ActiveEventSaisie(Status: boolean);
begin
  if not Status then
  begin
    TRadioButton(getControl('RPLUS')).OnClick := nil;
    TRadioButton(getControl('RMOINS')).OnClick := nil;
    THEdit(GetControl('BM6_ARTICLE')).OnElipsisClick := nil;
    THValComboBox (GetControl('BM6_TYPEERP')).OnChange := nil;
  end else
  begin
    TRadioButton(getControl('RPLUS')).OnClick := PlusMoinsClick;
    TRadioButton(getControl('RMOINS')).OnClick := PlusMoinsClick;
    THEdit(GetControl('BM6_ARTICLE')).OnElipsisClick := SearchRefErp;
    THValComboBox (GetControl('BM6_TYPEERP')).OnChange := ChangeType;
  end;
  //
end;

procedure TOF_BTPARAMLIGBAST.PrepareAffichage;
begin
  ListeChamps := 'BM6_INDICE;BM6_CODE;BM6_LIBELLE;';
  GS.RowCount := 2;
  GS.ColCount := 3;
  //
  GS.Cells[0,0] := ' ';
  GS.ColAligns[0] := taLeftJustify;
  GS.ColWidths[0] := 10*GS.Canvas.TextWidth('W');
  GS.Collengths[0] := 10*GS.Canvas.TextWidth('W');
  GS.ColEditables [0]:= false;
  //
  GS.Cells[1,0] := 'Code';
  GS.ColAligns[1] := taLeftJustify;
  GS.ColWidths[1] := 20*GS.Canvas.TextWidth('W');
  GS.Collengths[1] := 20*GS.Canvas.TextWidth('W');
  GS.ColEditables [1]:= false;
  //
  GS.Cells[2,0] := 'Définition';
  GS.ColAligns[2] := taLeftJustify;
  GS.ColWidths[2] := 70*GS.Canvas.TextWidth('W');
  GS.Collengths[2] := 70*GS.Canvas.TextWidth('W');
  GS.ColEditables [2]:= false;
end;

procedure TOF_BTPARAMLIGBAST.AfficheGrid;
var Ligne : Integer;
    TOBL : TOB;
begin
  GS.RowCount := TOBCat.Detail.Count +1;  if GS.rowCount < 2 then gs.rowCount := 2;
  GS.FixedCols := 1;
  GS.FixedRows := 1;
  for Ligne := 1 to TOBCat.Detail.Count do
  begin
    TOBL := GetTOBCat (TOBCat, Ligne); if TOBl = nil then exit;
    TOBL.PutLigneGrid (GS,Ligne,false,false,ListeChamps);
  end;
  HMTrad.ResizeGridColumns(GS);
end;

function TOF_BTPARAMLIGBAST.GetTOBCat(TOBCat: TOB; Ligne: integer): TOB;
begin
  Result := nil;
  if Ligne > TOBCat.detail.count then exit;
  Result := TOBCat.detail[Ligne-1];
end;

procedure TOF_BTPARAMLIGBAST.GSRowEnter(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
var TOBL : TOB;
    P1 : THPanel;
begin
  TOBL := GetTOBCat(TOBCat,Ou); if TOBL = nil then Exit;
  P1 := THpanel(GetControl('P1'));
  TOBL.PutEcran(Ecran,P1);
  if Pos(TOBL.GetString('BM6_TYPEERP'),'CAU;RG')>0 then
  begin
    THLabel(GetControl('LBM6_ARTICLE')).Visible := false;
    THEdit(GetControl('BM6_ARTICLE')).Visible := false;
  end else
  begin
    THLabel(GetControl('LBM6_ARTICLE')).Visible := true;
    THEdit(GetControl('BM6_ARTICLE')).Visible := true;
  end;
  if THEdit(GetControl('BM6_SENS')).Text = '+' then TRadioButton(getControl('RPLUS')).Checked := True
                                               else TRadioButton(getControl('RMOINS')).Checked := True;
  ActiveEventSaisie(true);
end;

procedure TOF_BTPARAMLIGBAST.GSRowExit(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
var TOBL : TOB;
    P1 : THPanel;
begin
  BeforeChange(Ou);
  ActiveEventSaisie(false);
end;

procedure TOF_BTPARAMLIGBAST.PlusMoinsClick(Sender: TObject);
begin
  if TRadioButton(getControl('RPLUS')).Checked then THEdit(GetControl('BM6_SENS')).Text := '+'
                                               else THEdit(GetControl('BM6_SENS')).Text := '-';
end;

procedure TOF_BTPARAMLIGBAST.SearchRefErp(Sender: TObject);
var CodeArt,StChamps : string;
    CODE: THCritMaskEdit;
begin
  StChamps := '';
  if THValComboBox (GetControl('BM6_TYPEERP')).Value = 'ART' then
  begin
    CodeArt := ThEdit(GetControl('BM6_ARTICLE')).Text;
    if CodeArt <> '' then StChamps := 'GA_CODEARTICLE=' + Trim(Copy(CodeArt, 1, 18)) + ';';
    if StChamps <> '' then stChamps := stChamps + ';';
    StChamps:= StChamps + 'XX_WHERE=GA_TYPEARTICLE="MAR" OR GA_TYPEARTICLE="ARP" OR GA_TYPEARTICLE="FRS"';
    //
    CodeArt := AGLLanceFiche('BTP', 'BTARTICLE_RECH', '', '',StChamps+';ACTION=CONSULTATION;STATUTART=UNI,DIM');

    if codeArt <> '' then
    begin
      ThEdit(GetControl('BM6_ARTICLE')).text := CodeArt;
    end;
  end else
  begin
    CODE          := THCritMaskEdit.Create(ECRAN);
    CODE.Parent   := ThEdit(GetControl('BM6_ARTICLE'));
    CODE.Top      := ThEdit(GetControl('BM6_ARTICLE')).Top;
    CODE.Left     := ThEdit(GetControl('BM6_ARTICLE')).Left;
    CODE.Width    := 3;
    CODE.Visible  := False;
    CODE.Text     := ThEdit(GetControl('BM6_ARTICLE')).Text;
    CODE.DataType := 'GCPORT';
    CODE.Plus     := ' AND (GPO_TYPEFRAIS <> "B00")';
    //
    LookupCombo(CODE);
    //
    if CODE.Text <> '' then ThEdit(GetControl('BM6_ARTICLE')).text := CODE.Text;
    //
    CODE.free;
  end;
end;

procedure TOF_BTPARAMLIGBAST.ChangeType(Sender: TObject);
begin
  if Pos(THValComboBox (GetControl('BM6_TYPEERP')).Value,'CAU;RG')>0 then
  begin
    THLabel(GetControl('LBM6_ARTICLE')).Visible := false;
    THEdit(GetControl('BM6_ARTICLE')).Visible := false;
  end else
  begin
    THLabel(GetControl('LBM6_ARTICLE')).Visible := true;
    THEdit(GetControl('BM6_ARTICLE')).Visible := true;
  end;
end;

procedure TOF_BTPARAMLIGBAST.PARAMBSVClick(Sender: TObject);
begin
  AglLanceFiche('BTP','BTPARPIECEBSV','XBT','','PIECE=XBT') ;
end;

Initialization
  registerclasses ( [ TOF_BTPARAMLIGBAST ] ) ;
end.
