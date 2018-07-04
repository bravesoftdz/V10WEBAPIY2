unit MulAna;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, Menus, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  Hqry, Grids, DBGrids, StdCtrls, Hctrls,
{$IFDEF COMPTA}
  SaisODA,
{$ENDIF}
  ExtCtrls, ComCtrls, Buttons, hmsgbox, Hcompte, Mask, HEnt1, SaisUtil, Ent1,
  SaisComm, HRichEdt, HSysMenu, HDB, HTB97, ColMemo, HPanel, UiUtil, CritEdt,
  Fe_Main,
  Filtre,
  HRichOLE, ADODB ;

procedure MultiCritereAna(Comment : TActionFiche);
procedure MultiCritereAnaANouveau(Comment : TActionFiche);
procedure MultiCritereAnaZoom(Comment : TActionFiche ; Var CritEdt : TCritEdt);

type
  TFMulAna = class(TFMul)
    TY_JOURNAL: TLabel;
    Y_JOURNAL: THValComboBox;
    TY_EXERCICE: THLabel;
    Y_EXERCICE: THValComboBox;
    TY_AXE: TLabel;
    Y_AXE: THValComboBox;
    Y_NUMVENTIL: THCritMaskEdit;
    Y_NUMVENTIL_: THCritMaskEdit;
    TY_NUMEROPIECE: TLabel;
    Y_NUMEROPIECE: THCritMaskEdit;
    HLabel1: THLabel;
    Y_NUMEROPIECE_: THCritMaskEdit;
    Y_DATECOMPTABLE_: THCritMaskEdit;
    TY_DATECOMPTABLE_: THLabel;
    Y_DATECOMPTABLE: THCritMaskEdit;
    TY_DATECOMPTABLE: THLabel;
    Y_TYPEANALYTIQUE: TCheckBox;
    Y_VALIDE: TCheckBox;
    TY_QUALIFPIECE: THLabel;
    Y_QUALIFPIECE: THValComboBox;
    TY_DATECREATION: THLabel;
    Y_DATECREATION: THCritMaskEdit;
    TY_DATECREATION_: THLabel;
    Y_DATECREATION_: THCritMaskEdit;
    Y_DEVISE: THValComboBox;
    TY_DEVISE: THLabel;
    RUne: TCheckBox;
    TY_GENERAL: TLabel;
    Y_GENERAL: THCpteEdit;
    TY_SECTION: TLabel;
    Y_SECTION: THCpteEdit;
    HM: THMsgBox;
    Pzlibre: TTabSheet;
    Bevel5: TBevel;
    TY_TABLE0: TLabel;
    TY_TABLE2: TLabel;
    TY_TABLE3: TLabel;
    TY_TABLE1: TLabel;
    Y_TABLE0: THCpteEdit;
    Y_TABLE2: THCpteEdit;
    Y_TABLE3: THCpteEdit;
    Y_TABLE1: THCpteEdit;
    TY_ETABLISSSEMENT: THLabel;
    Y_ETABLISSEMENT: THValComboBox;
    XX_WHERE: TEdit;
    BTobV: TToolbarButton97;
    procedure BChercheClick(Sender: TObject); override;
    procedure Y_EXERCICEChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Y_AXEChange(Sender: TObject);
    procedure RUneClick(Sender: TObject);
    procedure Y_QUALIFPIECEChange(Sender: TObject);
    procedure FListeDblClick(Sender: TObject); override;
    procedure HMTradBeforeTraduc(Sender: TObject);
    procedure FormCreate(Sender: TObject); override;
    procedure BTobVClick(Sender: TObject);
  private    { Déclarations privées }
    ANouveau : Boolean ;
    OkZoom : Boolean ;
    CritEdt : TCritEdt ;
    procedure RetoucheCritere ;
    procedure InitCriteres ;
    procedure ErgoS3 ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  ULibExercice,
  CPTypeCons,
  CPProcMetier,
  {$ENDIF MODENT1}
  ParamSoc {SG6 23/12/2004 Gestion du mdoe croisaxe};

procedure MultiCritereAna(Comment : TActionFiche);
var FMulAna : TFMulAna ;
    M       : RMVT ;
    PP : THPanel ;
begin
if Comment=taCreat then
   BEGIN
   FillChar(M,Sizeof(M),#0) ; M.Simul:='N' ; M.CodeD:=V_PGI.DevisePivot ;
   M.DateC:=V_PGI.DateEntree ; M.TauxD:=1 ; M.DateTaux:=M.DateC ; M.Valide:=False ;
   M.Etabl:=VH^.ETABLISDEFAUT ;
{$IFDEF COMPTA}
   LanceSaisieODA(Nil,taCreat,M) ;
{$ENDIF}
   Exit ;
   END ;
FMulAna:=TFMulAna.Create(Application) ;
FMulAna.TypeAction:=Comment ;
FMulAna.Q.Manuel:=TRUE ;
FMulAna.FNomFiltre:='MULANA'+Chr(48+Ord(Comment)) ;
FMulAna.ANouveau:=FALSE ; FMulAna.OkZoom:=FALSE ;
Case Comment of
  taConsult : begin
              FMulAna.Caption:=FMulAna.HM.Mess[0] ;
              FMulAna.Q.Liste:='MULVANAL' ;
              end ;
  taModif   : begin
              FMulAna.Caption:=FMulAna.HM.Mess[1] ;
              FMulAna.Q.Liste:='MULMANAL' ;
              end ;
  end ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FMulAna.ShowModal ;
    finally
     FMulAna.Free ;
    end;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FMulAna,PP) ;
   FMulAna.Show ;
   END ;
end;

procedure MultiCritereAnaANouveau(Comment : TActionFiche);
var FMulAna : TFMulAna ;
    M       : RMVT ;
    PP : THPanel ;
begin
if Comment=taCreat then
   BEGIN
   FillChar(M,Sizeof(M),#0) ; M.Simul:='N' ; M.CodeD:=V_PGI.DevisePivot ;
   M.DateC:=V_PGI.DateEntree ; M.TauxD:=1 ; M.DateTaux:=M.DateC ; M.Valide:=False ;
   M.Etabl:=VH^.ETABLISDEFAUT ; M.Anouveau:=TRUE ;
{$IFDEF COMPTA}
   LanceSaisieODA(Nil,taCreat,M) ;
{$ENDIF}
   Exit ;
   END ;
FMulAna:=TFMulAna.Create(Application) ;
FMulAna.TypeAction:=Comment ;
FMulAna.Q.Manuel:=TRUE ;
FMulAna.ANouveau:=TRUE ; FMulAna.OkZoom:=FALSE ;
FMulAna.FNomFiltre:='MULANAAN'+Chr(48+Ord(Comment)) ;
Case Comment of
  taConsult : begin
              FMulAna.Caption:=FMulAna.HM.Mess[0] ;
              If FMulAna.ANouveau Then FMulAna.Caption:=FMulAna.HM.Mess[3] ;
              FMulAna.Q.Liste:='MULVANAL' ;
              end ;
  taModif   : begin
              FMulAna.Caption:=FMulAna.HM.Mess[1] ;
              If FMulAna.ANouveau Then FMulAna.Caption:=FMulAna.HM.Mess[4] ;
              FMulAna.Q.Liste:='MULMANAL' ;
              end ;
  end ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FMulAna.ShowModal ;
    finally
     FMulAna.Free ;
    end;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FMulAna,PP) ;
   FMulAna.Show ;
   END ;
end;

procedure MultiCritereAnaZoom(Comment : TActionFiche ; Var CritEdt : TCritEdt);
var FMulAna : TFMulAna ;
    M       : RMVT ;
    PP : THPanel ;
begin
if Comment=taCreat then
   BEGIN
   FillChar(M,Sizeof(M),#0) ; M.Simul:='N' ; M.CodeD:=V_PGI.DevisePivot ;
   M.DateC:=V_PGI.DateEntree ; M.TauxD:=1 ; M.DateTaux:=M.DateC ; M.Valide:=False ;
   M.Etabl:=VH^.ETABLISDEFAUT ;
{$IFDEF COMPTA}
   LanceSaisieODA(Nil,taCreat,M) ;
{$ENDIF}
   Exit ;
   END ;
FMulAna:=TFMulAna.Create(Application) ;
FMulAna.TypeAction:=Comment ;
FMulAna.Q.Manuel:=TRUE ;
FMulAna.FNomFiltre:='MULANA'+Chr(48+Ord(Comment)) ;
FMulAna.ANouveau:=FALSE ; FMulAna.OkZoom:=TRUE ; FMulAna.CritEdt:=CritEdt ;
Case Comment of
  taConsult : begin
              FMulAna.Caption:=FMulAna.HM.Mess[0] ;
              FMulAna.Q.Liste:='MULVANAL' ;
              end ;
  taModif   : begin
              FMulAna.Caption:=FMulAna.HM.Mess[1] ;
              FMulAna.Q.Liste:='MULMANAL' ;
              end ;
  end ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FMulAna.ShowModal ;
    finally
     FMulAna.Free ;
    end;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FMulAna,PP) ;
   FMulAna.Show ;
   END ;
end;



procedure TFMulAna.BChercheClick(Sender: TObject);
begin
RetoucheCritere ;
  inherited;
end;

procedure TFMulAna.Y_EXERCICEChange(Sender: TObject);
begin
  inherited;
ExoToDates(Y_EXERCICE.Value,Y_DATECOMPTABLE,Y_DATECOMPTABLE_) ;
end;

procedure TFMulAna.ErgoS3 ;
BEGIN
if Not EstSerie(S3) then Exit ;
Y_TYPEANALYTIQUE.Visible:=False ;
Y_AXE.Visible:=False ; TY_AXE.Visible:=False ; 
END ;

procedure TFMulAna.FormShow(Sender: TObject);
Var Cpt1 : String17 ;
    Date1,Date2 : TDateTime ;
begin
if ANouveau then
  begin
    case TypeAction of
      taCreat   : HelpContext:=7728000 ;
      taConsult : HelpContext:=7729000 ;
       taModif  : HelpContext:=7730000 ;
    end ;
  end else
  begin
    case TypeAction of
      taConsult : HelpContext:=7364000 ;
       taModif : HelpContext:=7370000 ;
    end ;
  end ;
InitCriteres ;
Cpt1:=SQLPremierDernier(fbJal,TRUE) ;
if ((OkZoom) and (CritEDT.sCpt1<>'')) then else Y_Journal.Value:=Cpt1 ;
Y_EXERCICE.Value:=QuelDateExo(V_PGI.DateEntree,Date1,Date2) ;
If ANouveau Then
  BEGIN
  Y_DATECOMPTABLE.Text:=DateToStr(VH^.Entree.Deb) ;
  Y_DATECOMPTABLE_.Text:=DateToStr(VH^.Entree.Deb) ;
  END Else
  BEGIN
  Y_DATECOMPTABLE.Text:=DateToStr(V_PGI.DateEntree) ;
  Y_DATECOMPTABLE_.Text:=DateToStr(V_PGI.DateEntree) ;
  END ;
Y_DATECREATION.Text:=StDate1900 ; Y_DATECREATION_.Text:=StDate2099 ;
Y_AXE.Value:='A1' ;
PositionneEtabUser(Y_ETABLISSEMENT) ;
if ANouveau And (Y_JOURNAL.Items.Count>0) then
   BEGIN
   if Y_JOURNAL.Vide then Y_JOURNAL.ItemIndex:=1 else Y_JOURNAL.ItemIndex:=0 ;
   END ;
If OkZoom Then
  BEGIN
  if CritEDT.Exo.Code<>'' then Y_EXERCICE.Value:=CritEdt.Exo.Code ;
  Y_DATECOMPTABLE.Text:=DateToStr(CritEdt.Date1) ;
  Y_DATECOMPTABLE_.Text:=DateToStr(CritEdt.Date2) ;
  if CritEdt.Cpt1<>'' then
     BEGIN
     Y_SECTION.Text:=CritEdt.Cpt1 ;
     Y_AXE.Value:=CritEdt.Bal.Axe ;
     Y_JOURNAL.ItemIndex:=0 ;
     END else
     BEGIN
     Y_GENERAL.Text:=CritEdt.sCpt1 ;
     END ;
  END ;
  inherited;
Q.Manuel:=FALSE ;
Q.UpdateCriteres ;
CentreDBGrid(FListe) ;
ErgoS3 ;
{$IFDEF CCSTD}
  ListeFiltre.ForceAccessibilite := FaPublic;
{$ENDIF}
HMTrad.ResizeDBGridColumns(FListe) ;
end;

procedure TFMulAna.RetoucheCritere ;
begin
If (Y_GENERAL.Text<>'') Or (Y_SECTION.Text<>'') then RUne.Checked:=FALSE ;
end ;

procedure TFMulAna.Y_AXEChange(Sender: TObject);
Var tt : TZoomTable ;
    i : Byte ;
begin
  inherited;
if Y_AXE.Value='' then Exit ;
tt:=tzSection ; Dec(tt) ; i:=StrToInt(Copy(Y_AXE.Value,2,1)) ; Byte(tt):=Byte(tt)+i ;
Y_SECTION.ZoomTable:=tt ;
end;

procedure TFMulAna.InitCriteres ;
BEGIN
if VH^.Precedent.Code<>'' then Y_DATECOMPTABLE.Text:=DateToStr(VH^.Precedent.Deb)
                          else Y_DATECOMPTABLE.Text:=DateToStr(VH^.Encours.Deb) ;
Y_DATECOMPTABLE_.Text:=DateToStr(V_PGI.DateEntree) ;
If ANouveau Then
  BEGIN
  Y_JOURNAL.DataType:='TTJALANALAN' ; XX_WHERE.Text:='Y_ECRANOUVEAU="OAN"' ;
  Y_TYPEANALYTIQUE.Enabled:=FALSE ;
  TY_QUALIFPIECE.Enabled:=FALSE ; Y_QUALIFPIECE.Enabled:=FALSE ; Y_QUALIFPIECE.value := 'N' ; 
  TY_DEVISE.Enabled:=FALSE ; Y_DEVISE.Enabled:=FALSE ; Y_DEVISE.ItemIndex:=0 ;
  END Else XX_WHERE.Text:='Y_ECRANOUVEAU<>"OAN"' ;
END ;

procedure TFMulAna.RUneClick(Sender: TObject);
begin
  inherited;
If RUne.Checked Then BEGIN Y_NUMVENTIL.text:='1' ; Y_NUMVENTIL_.text:='1' ; END
                Else BEGIN Y_NUMVENTIL.text:='0' ; Y_NUMVENTIL_.text:='9999' ; END ;
end;

procedure TFMulAna.Y_QUALIFPIECEChange(Sender: TObject);
begin
  inherited;
if Y_QUALIFPIECE.Value<>'R' then Exit ;
if Not V_PGI.Controleur then Y_QUALIFPIECE.Value:='N' ;
end;

procedure TFMulAna.FListeDblClick(Sender: TObject);
begin
  inherited;
if(Q.Eof)And(Q.Bof) then Exit ;
{$IFDEF COMPTA}
TrouveEtLanceSaisieODA(Q,TypeAction) ;
{$ENDIF}
end;

procedure TFMulAna.HMTradBeforeTraduc(Sender: TObject);
begin
  inherited;
LibellesTableLibre(PzLibre,'TY_TABLE','Y_TABLE','Y') ;
end;

procedure TFMulAna.FormCreate(Sender: TObject);
begin
  inherited;
  //SG6 23/12/2004 mode croisaxe
  if VH^.AnaCroisaxe then
  begin
    Y_SECTION.Visible := False;
    TY_SECTION.Visible := False;
    Y_AXE.Visible := False;
    TY_AXE.Visible := False;
  end;
  MemoStyle:=msBook ;
end;

procedure TFMulAna.BTobVClick(Sender: TObject);
Var WhereSQL : String ;
begin
  inherited;
if ((Y_JOURNAL.Value='') and (Y_GENERAL.Text='') and (Y_SECTION.Text='')) then
   if HM.Execute(5,caption,'')<>mrYes then Exit ;
WhereSQL:=RecupWhereCritere(Pages) ;
AGLLanceFiche('CP','CPCONSULTREVIS','','',WhereSQL) ;
end;

end.
