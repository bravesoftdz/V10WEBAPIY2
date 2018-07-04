unit FmtChoix;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, Ent1,
  StdCtrls, Buttons, ExtCtrls, ComCtrls, Hctrls, DB, {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} hmsgbox, HEnt1,
  HSysMenu, HTB97, Hqry, HPanel, UiUtil,RappType, ParamSoc, ADODB ;

type TFmtChoix = RECORD
    OkSauve : boolean ;
    Format : String3 ;
    FixeFmt : Boolean ;
    Fichier : String ;
    FixeFichier : Boolean ;
    CompteRendu : boolean ;
    Cycle : boolean ;
    Ascii : Boolean ;
    Detruire : Boolean ;
    Valide : Boolean ;
    END ;

var ChoixFmt : TFmtChoix ;

type
  TFFmtChoix = class(TForm)
    Pages: TTabControl;
    Panel2: TPanel;
    GroupBox1: TGroupBox;
    FFmt: THValComboBox;
    Format: THLabel;
    FFile: TEdit;
    HLabel2: THLabel;
    TFMTCHOIX: THTable;
    MsgBox: THMsgBox;
    FFixeFmt: TCheckBox;
    FFixeFile: TCheckBox;
    BNomFichier: TButton;
    SaveFile: TSaveDialog;
    TFMTCHOIXIC_IMPORT: TStringField;
    TFMTCHOIXIC_NATURE: TStringField;
    TFMTCHOIXIC_FORMAT: TStringField;
    TFMTCHOIXIC_DETRUIRE: TStringField;
    TFMTCHOIXIC_RESUME: TStringField;
    TFMTCHOIXIC_VALIDE: TStringField;
    TFMTCHOIXIC_RECYCLE: TStringField;
    TFMTCHOIXIC_FICHIER: TStringField;
    TFMTCHOIXIC_ASCII: TStringField;
    TFMTCHOIXIC_FIXEFORMAT: TStringField;
    TFMTCHOIXIC_FIXEFICHIER: TStringField;
    GOptions: TGroupBox;
    FRes: TCheckBox;
    FASCII: TCheckBox;
    FDelImp: TCheckBox;
    FValidImp: TCheckBox;
    Label1: TLabel;
    HMTrad: THSystemMenu;
    Dock971: TDock97;
    Panel1: TToolWindow97;
    BParam: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    BScenario: TToolbarButton97;
    BFourCpt: TToolbarButton97;
    procedure PagesChanging(Sender: TObject; var AllowChange: Boolean);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure Modif(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure PagesChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BNomFichierClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BParamClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BScenarioClick(Sender: TObject);
    procedure FFmtChange(Sender: TObject);
    procedure BFourCptClick(Sender: TObject);
  private    { Déclarations privées }
    EnModif : Boolean ;
    Imp,CurNat : String ;
    Modifie : Boolean ;
    Function  Sauve(Question : boolean) : Boolean ;
    procedure Charge ;
    procedure AjouteTabs ;
    procedure ModifFmtChoixTmp ;
    procedure ResizeEcran ;
    procedure SauveEtapes ;
    procedure ScenarioVisible ;
  public
    { Déclarations publiques }
  end;

Procedure ChoixFormatImpExp(Imp : String ; CurNat : String) ;
procedure ChangeDataTypeFmt(Imp : String ; CurNat : String ; FFmt : THValComboBox) ;
procedure GereFMTChoix(Imp,CurNat : String ; FFormat : THValComboBox ;FileName : TEdit ;
                       RechFile : TToolbarButton97) ;
function  FormatOk(NomFic : String ; Fmt,Lequel : String3) : boolean ;

implementation

uses
  {$IFDEF MODENT1}
  CPProcGen,
  {$ENDIF MODENT1}
  FmtImpor,ScenaCom,ScenaCoE
  {$IFNDEF TOPGI}
  ,FouCpt
  {$ENDIF}
  ;

{$R *.DFM}

Procedure ChoixFormatImpExp(Imp : String ; CurNat : String) ;
var FFmtChoix: TFFmtChoix;
    PP : THPanel ;
BEGIN
FFmtChoix:=TFFmtChoix.Create(Application) ;
FFmtChoix.Imp:=Imp ; if Imp='-' then FFmtChoix.HelpContext:=6120000 ;
FFmtChoix.CurNat:=CurNat ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
   try
     FFmtChoix.ShowModal ;
     Finally
     FFmtChoix.Free ;
     END ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FFmtChoix,PP) ;
   FFmtChoix.Show ;
   END ;
END ;

procedure TFFmtChoix.FormShow(Sender: TObject);
var Mss : integer ;
BEGIN
Mss:=12 ;
if Imp='-' then Mss:=Mss+11 ;
ResizeEcran ;
AjouteTabs ;
EnModif :=False ;
if CurNat='' then BEGIN EnModif :=True ; CurNat:='FGE' END else
  BEGIN
  if CurNat='FGE' then Mss:=Mss+1 else
  if CurNat='FSE' then Mss:=Mss+2 else
  if CurNat='FTI' then Mss:=Mss+3 else
  if CurNat='FBJ' then Mss:=Mss+4 else
  if CurNat='FBG' then Mss:=Mss+5 else
  if CurNat='FBS' then Mss:=Mss+6 else
  if CurNat='FEC' then Mss:=Mss+7 else
  if CurNat='FOD' then Mss:=Mss+8 else
  if CurNat='FBE' then Mss:=Mss+9 else
  if CurNat='FBA' then Mss:=Mss+10 else
  If CurNat='FCC' Then BEGIN If Imp='-' Then Mss:=37 ; END ;
  END ;
Caption:=Msgbox.Mess[Mss] ;
ChangeDataTypeFmt(Imp,CurNat,FFmt) ;
ScenarioVisible ;
Pages.TabIndex:=0 ;
Charge ;
UpdateCaption(Self) ;
If Imp='X' Then
   BEGIN
   BScenario.Caption:=MsgBox.Mess[35] ;
   BScenario.Hint:=MsgBox.Mess[35] ;
   END Else
   BEGIN
   BScenario.Caption:=MsgBox.Mess[36] ;
   BScenario.Hint:=MsgBox.Mess[36] ;
   END ;
END ;

procedure TFFmtChoix.ResizeEcran ;
BEGIN
if Imp='-' then
  BEGIN
  if not FDelImp.Visible then exit ;
  FASCII.Caption:=Msgbox.Mess[1] ;
  FDelImp.Visible:=False ;
  FValidImp.Visible:=False ;
  GOptions.Height:=GOptions.Height-65 ;
  ClientHeight:=ClientHeight-65 ;
  END else
  if (CurNat<>'FEC') and (CurNat<>'FBA')
   and  (CurNat<>'FOD') and (CurNat<>'FBE')then
    BEGIN
    if not FValidImp.Visible then exit ;
    FValidImp.Visible:=False ;
    GOptions.Height:=GOptions.Height-25 ;
    ClientHeight:=ClientHeight-25 ;
    END else
    BEGIN
    if FValidImp.Visible then exit ;
    FValidImp.Visible:=True ;
    GOptions.Height:=GOptions.Height+25 ;
    ClientHeight:=ClientHeight+25 ;
    END ;
END ;

procedure TFFmtChoix.Charge ;
BEGIN
TFMTCHOIX.Open ;
if TFMTCHOIX.FindKey([Imp,CurNat]) then
   BEGIN
   FFmt.Value:=TFMTCHOIX.FindField('IC_FORMAT').AsString ;
   FDelImp.Checked:=(TFMTCHOIX.FindField('IC_DETRUIRE').AsString='X') ;
   FRes.Checked:=(TFMTCHOIX.FindField('IC_RESUME').AsString='X') ;
   //FCycle.Checked:=(TFMTCHOIX.FindField('IC_RECYCLE').AsString='X') ;
   FValidImp.Checked:=(TFMTCHOIX.FindField('IC_VALIDE').AsString='X') ;
   FFile.Text:=TFMTCHOIX.FindField('IC_FICHIER').AsString ;
   FASCII.Checked:=(TFMTCHOIX.FindField('IC_ASCII').AsString='X') ;
   FFixeFmt.Checked:=(TFMTCHOIX.FindField('IC_FIXEFORMAT').AsString='X') ; ;
   FFixeFile.Checked:=(TFMTCHOIX.FindField('IC_FIXEFICHIER').AsString='X') ; ;
   END else
   BEGIN
   FFmt.ItemIndex:=-1 ;
   FDelImp.Checked:=False ;
   FRes.Checked:=False ;
   //FCycle.Checked:=False ;
   FValidImp.Checked:=False ;
   FFile.Text:='' ;
   FASCII.Checked:=False ;
   FFixeFmt.Checked:=False ;
   FFixeFile.Checked:=FAlse ;
   END ;
TFMTCHOIX.Close ;
Modifie:=FALSE ;
END ;

Function TFFmtChoix.Sauve(Question : boolean) : boolean ;
var rep : integer ;
BEGIN
Sauve:=TRUE ;
if Not Modifie then Exit ;
if Question then Rep:=MsgBox.execute(0,caption,'') else Rep:=mrYes ;
Case rep of
  mrYes : BEGIN
          TFMTCHOIX.Open ;
          if TFMTCHOIX.FindKey([Imp,CurNat]) then TFMTCHOIX.Edit
            else
            BEGIN
            TFMTCHOIX.Insert ;
            TFMTCHOIX.FindField('IC_IMPORT').AsString:=Imp ;
            TFMTCHOIX.FindField('IC_NATURE').AsString:=CurNat ;
            END ;
          TFMTCHOIX.FindField('IC_FORMAT').AsString:=FFmt.Value ;
          TFMTCHOIX.FindField('IC_FICHIER').AsString:=FFile.Text ;
          if FDelImp.Checked then TFMTCHOIX.FindField('IC_DETRUIRE').AsString:='X'
            else TFMTCHOIX.FindField('IC_DETRUIRE').AsString:='-' ;
          if FRes.Checked then TFMTCHOIX.FindField('IC_RESUME').AsString:='X'
            else TFMTCHOIX.FindField('IC_RESUME').AsString:='-' ;
          //if FCycle.Checked then TFMTCHOIX.FindField('IC_RECYCLE').AsString:='X'
          //  else TFMTCHOIX.FindField('IC_RECYCLE').AsString:='-' ;
          if FValidImp.Checked then TFMTCHOIX.FindField('IC_VALIDE').AsString:='X'
            else TFMTCHOIX.FindField('IC_VALIDE').AsString:='-' ;
          if FASCII.Checked then TFMTCHOIX.FindField('IC_ASCII').AsString:='X'
            else TFMTCHOIX.FindField('IC_ASCII').AsString:='-' ;
          if FFixeFmt.Checked then TFMTCHOIX.FindField('IC_FIXEFORMAT').AsString:='X'
            else TFMTCHOIX.FindField('IC_FIXEFORMAT').AsString:='-' ;
          if FFixeFile.Checked then TFMTCHOIX.FindField('IC_FIXEFICHIER').AsString:='X'
            else TFMTCHOIX.FindField('IC_FIXEFICHIER').AsString:='-' ;
          TFMTCHOIX.Post ;
          TFMTCHOIX.Close ;
          Modifie:=FALSE ;
          if (CurNat='FEC') or (CurNat='FBA') then SauveEtapes ;
          END  ;
  mrNo : ;
  mrCancel : Sauve:=FALSE ;
  end ;

END ;

procedure TFFmtChoix.SauveEtapes ;
var s   : String ;
    i,D : integer ;
{$IFDEF SPEC302}
    Q : TQuery ;
{$ENDIF}
BEGIN
s := '' ;
{$IFDEF SPEC302}
Q:=OpenSQL('SELECT SO_ETAPEIMPORT FROM SOCIETE',True) ;
if not Q.Eof then s:=Q.Fields[0].AsString ;
Ferme(Q) ;
{$ELSE}
s:=GetParamSocSecur('SO_ETAPEIMPORT', False) ;
{$ENDIF}
if s='' then while Length(s)<=20 do s:=s+'0' ;
D:=1 ; if CurNat='FBA' then D:=4 ;
for i:=D to D+3 do s[i]:='0' ;
{$IFDEF SPEC302}
ExecuteSQL('UPDATE SOCIETE SET SO_ETAPEIMPORT="'+s+'"') ;
{$ELSE}
SetParamSoc('SO_ETAPEIMPORT',s) ;
{$ENDIF}
END ;

procedure TFFmtChoix.PagesChanging(Sender: TObject; var AllowChange: Boolean);
begin
AllowChange:=Sauve(TRUE) ;
end;

procedure TFFmtChoix.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
if EnModif then CanClose:=Sauve(TRUE) ;
end;

procedure TFFmtChoix.Modif(Sender: TObject);
begin
Modifie:=TRUE ;
end;

procedure TFFmtChoix.BValiderClick(Sender: TObject);
begin
if EnModif then Sauve(FALSE) else BEGIN Modifie:=FALSE ; ModifFmtChoixTmp ; END ;

Close;
if isInside(self) then begin
  THPanel(parent).CloseInside;
  THPanel(parent).VideToolBar;
end;
end;

procedure TFFmtChoix.BFermeClick(Sender: TObject);
begin
Close;
if isInside(self) then begin
  THPanel(parent).CloseInside;
  THPanel(parent).VideToolBar;
end;
end;

procedure TFFmtChoix.PagesChange(Sender: TObject);
begin
Case Pages.TabIndex of
   0 : CurNat:='FGE' ; { Comptes généraux }
   1 : CurNat:='FTI' ; { Comptes auxiliaires }
   2 : CurNat:='FSE' ; { Comptes analytiques }
   3 : CurNat:='FBJ' ; { Journaux }
   4 : CurNat:='FBG' ; { Comptes budgétaires }
   5 : CurNat:='FBS' ; { Sections budgétaires }
   6 : CurNat:='FEC' ; { Ecriture générale }
   7 : CurNat:='FOD' ; { Ecriture analytique }
   8 : CurNat:='FBE' ; { Ecriture budgétaire }
   9 : CurNat:='FBA' ; { Balance }
   10 : CurNat:='FCC' ; { En cours }
   END ;
BParam.Enabled:=(CurNat<>'FBA') ;
ChangeDataTypeFmt(Imp,CurNat,FFmt) ;
ScenarioVisible ;
Charge ;
ResizeEcran ;
end;

procedure TFFmtChoix.AjouteTabs ;
var i : integer ;
BEGIN
Pages.Tabs.Clear ;
if CurNat='' then
  BEGIN
  for i:=2 to 11 do Pages.Tabs.Add(Msgbox.Mess[i]) ;
  If Imp='-' Then Pages.Tabs.Add(Msgbox.Mess[38]) ;
  END else
  BEGIN
  if CurNat='FGE' then Pages.Tabs.Add(Msgbox.Mess[2]) else
  if CurNat='FTI' then Pages.Tabs.Add(Msgbox.Mess[3]) else
  if CurNat='FSE' then Pages.Tabs.Add(Msgbox.Mess[4]) else
  if CurNat='FBJ' then Pages.Tabs.Add(Msgbox.Mess[5]) else
  if CurNat='FBG' then Pages.Tabs.Add(Msgbox.Mess[6]) else
  if CurNat='FBS' then Pages.Tabs.Add(Msgbox.Mess[7]) else
  if CurNat='FEC' then Pages.Tabs.Add(Msgbox.Mess[8]) else
  if CurNat='FOD' then Pages.Tabs.Add(Msgbox.Mess[9]) else
  if CurNat='FBE' then Pages.Tabs.Add(Msgbox.Mess[10]) else
  if CurNat='FBA' then Pages.Tabs.Add(Msgbox.Mess[11]) else
  if CurNat='FCC' then BEGIN If Imp='-' Then Pages.Tabs.Add(Msgbox.Mess[38]) ; END ; 
  END ;

END ;

procedure TFFmtChoix.BNomFichierClick(Sender: TObject);
begin
DirDefault(SaveFile,FFile.Text) ;
if SaveFile.Execute then FFile.Text:=SaveFile.FileName ;
end;

procedure AjouteParamFormat(Imp : String ; CurNat:  String ; var It,Va : TStrings) ;
var VaP,ItP : TStrings ;
    tt : string ;
BEGIN
VaP:=TStringList.Create ; ItP:=TStringList.Create ;
tt:='' ;
if CurNat='FGE' then BEGIN if Imp='-' then tt:='ttExpParamGen' else tt:='ttImpParamGen' ; END else
if CurNat='FTI' then BEGIN if Imp='-' then tt:='ttExpParamTie' else tt:='ttImpParamTie' ; END else
if CurNat='FSE' then BEGIN if Imp='-' then tt:='ttExpParamSec' else tt:='ttImpParamSec' ; END else
if CurNat='FBJ' then BEGIN if Imp='-' then tt:='ttExpParamJB' else tt:='ttImpParamJB' ; END else
if CurNat='FBG' then BEGIN if Imp='-' then tt:='ttExpParamCB' else tt:='ttImpParamCB' ; END else
if CurNat='FBS' then BEGIN if Imp='-' then tt:='ttExpParamSB' else tt:='ttImpParamSB' ; END else
if CurNat='FEC' then BEGIN if Imp='-' then tt:='ttExpParamEcr' else tt:='ttImpParamEcr' ; END else
if CurNat='FOD' then BEGIN if Imp='-' then tt:='ttExpParamOda' else tt:='ttImpParamOda' ; END else
if CurNat='FCC' then BEGIN if Imp='-' then tt:='ttExpParamEcc' else tt:='' ; END else
if CurNat='FBE' then BEGIN if Imp='-' then tt:='ttExpParamBud' else tt:='ttImpParamBud' ; END ;
//if CurNat='FBA' then BEGIN if Imp='-' then tt:=ttExpParamBal else tt:=ttImpParamBal ; END ;
if tt<>'' then RemplirValCombo(tt,'','',ItP,VaP,False,False) ;
if not (((ItP.Count=1) and (Trim(ItP[0])='')) or (ItP.Count=0)) then It.AddStrings(ItP) ;
if not (((VaP.Count=1) and (Trim(VaP[0])=''))or (VaP.Count=0)) then Va.AddStrings(VaP) ;
VaP.Clear ; VaP.Free ; ItP.Clear ; ItP.Free ;
END ;

procedure TFFmtChoix.ScenarioVisible ;
Var OkOk : Boolean ;
BEGIN
OkOk:=(CurNat='FEC') And (Imp='X') And
      ((FFmt.Value='SN2') Or (FFmt.Value='SAA') Or (FFmt.Value='HAL') Or
      (FFmt.Value='HLI') Or (FFmt.Value='CGN') Or (FFmt.Value='CGE') Or (FFmt.Value='SIS')) ;
OkOk:=OkOk Or ((CurNat='FEC') And (Imp='-')) ;
BScenario.Visible:=OkOk ;
OkOk:=(CurNat='FGE') And (Imp='X') And ((FFmt.Value='CGN') Or (FFmt.Value='CGE')) ;
BFourCpt.Visible:=OkOk ;
END ;

procedure ChangeDataTypeFmt(Imp : String ; CurNat : String ; FFmt : THValComboBox) ;
var Va,It : TStrings ;
BEGIN
Va:=TStringList.Create ; It:=TStringList.Create ;
if CurNat='FGE' then BEGIN If Imp='X' Then FFmt.DataType:='ttFormatGeneraux' Else FFmt.DataType:='ttFormatGenerauxExp'; END else
if CurNat='FTI' then BEGIN If Imp='X' Then FFmt.DataType:='ttFormatTiers' Else FFmt.DataType:='ttFormatTiersExp'  ; END else
if CurNat='FSE' then BEGIN If Imp='X' Then FFmt.DataType:='ttFormatSection' Else FFmt.DataType:='ttFormatSectionExp' ; END else
if CurNat='FBJ' then BEGIN FFmt.DataType:='' ; END else
if CurNat='FBG' then BEGIN FFmt.DataType:='' ; END else
if CurNat='FBS' then BEGIN FFmt.DataType:='' ; END else
if CurNat='FEC' then BEGIN If Imp='X' Then FFmt.DataType:='ttFormatEcriture' Else FFmt.DataType:='ttFormatEcritureExp' ; END else
if CurNat='FOD' then BEGIN If Imp='X' Then FFmt.DataType:='ttFormatEcrAna' Else FFmt.DataType:='ttFormatEcrAnaExp' ; END else
if CurNat='FBE' then BEGIN If Imp='X' Then FFmt.DataType:='ttFormatEcrBud' Else FFmt.DataType:='ttFormatEcrBudExp' ; END else
if CurNat='FCC' then BEGIN If Imp='X' Then FFmt.DataType:='ttFormatEcrBud' Else FFmt.DataType:='ttFormatEnCoursExp' ; END else
if CurNat='FBA' then BEGIN If Imp='X' Then FFmt.DataType:='ttFormatBalance' Else FFmt.DataType:='ttFormatBalanceExp' ; END ;
if FFmt.DataType<>'' then RemplirValCombo(FFmt.DataType,'','',It,Va,False,False) ;
//Obligé à cause du remplirliste qui ajoute un ' '...
if (((It.Count=1) and (Trim(It[0])='')) or (It.Count=0)) then It.Clear ;
if (((Va.Count=1) and (Trim(Va[0])=''))or (Va.Count=0)) then Va.Clear ;
AjouteParamFormat(Imp,CurNat,It,Va) ;
FFmt.DataType:='' ; FFmt.Values.Assign(Va) ; FFmt.Items.Assign(It) ;
Va.Clear ; Va.Free ; It.Clear ; It.Free ;
END ;

procedure GereFMTChoix(Imp,CurNat : String ; FFormat : THValComboBox ;FileName : TEdit ; RechFile : TToolbarButton97) ;
Var Q : TQuery ;
    FmtFixe, FixeFile : Boolean ;
BEGIN
With ChoixFmt do
  BEGIN
  OkSauve:=False ;
  Format:='' ;
  FixeFmt:=False ;
  Fichier:='' ;
  FixeFichier:=False ;
  END ;
Q:=OpenSQL('Select * FROM FMTCHOIX WHERE IC_IMPORT="'+Imp+'"'+
           ' AND IC_NATURE="'+CurNat+'"',True) ;
if Q.EOF then BEGIN Ferme(Q) ; Exit ; END ;
FileName.Text:=Q.FindField('IC_FICHIER').AsString ;
FmtFixe:=(Q.FindField('IC_FIXEFORMAT').AsString='X') ;
FixeFile:=(Q.FindField('IC_FIXEFICHIER').AsString='X') ;
FFormat.Value:=Q.FindField('IC_FORMAT').AsString ;
FFormat.Enabled:=not FmtFixe ;
FileName.Enabled:=not FixeFile ; RechFile.Enabled:=not FixeFile ;
With ChoixFmt do
  BEGIN
  CompteRendu:=(Q.FindField('IC_RESUME').AsString='X') ;
  Ascii:=(Q.FindField('IC_ASCII').AsString='X') ;
  Detruire:=(Q.FindField('IC_DETRUIRE').AsString='X') ;
  Valide:=(Q.FindField('IC_VALIDE').AsString='X') ;
  END ;
Ferme(Q) ;
END ;

procedure TFFmtChoix.ModifFmtChoixTmp ;
BEGIN
With ChoixFmt do
  BEGIN
  OkSauve :=True ;
  Format:=FFmt.Value ;
  FixeFmt:=(FFixeFmt.checked) ;
  Fichier :=FFile.Text ;
  FixeFichier:=(FFixeFile.checked) ;
  CompteRendu:=(FRes.checked) ;
  Ascii:=(FASCII.checked) ;
  Cycle:=False ; //(FCycle.checked) ;
  Detruire:=(FDelImp.checked) ;
  Valide:=(FValidImp.checked) ;
  END ;
END ;


function FormatOk(NomFic : String ; Fmt,Lequel : String3) : boolean ;
var F: TextFile ;
    St : String ;
    FormatFichier,TypeExp : String3 ;
BEGIN
if (Lequel='FBG') or (Lequel='FBS') or (Lequel='FBJ') or (Fmt='XRL') then BEGIN Result:=True ; Exit ; END ;
If (Fmt='CGN') And ((Lequel='X') Or (Lequel='T') Or (Lequel='G') Or (Lequel='A')) Then BEGIN Result:=True ; Exit ; END ;
Result:=False ;
AssignFile(F,NomFic) ;
{$I-} Reset (F,NomFic) ; {$I+}
if IoResult<>0 then Exit ;
Readln(F,St) ;
FormatFichier:=Trim(Copy(St,71,3)) ;
TypeExp:=Trim(Copy(St,74,3)) ;
if (FormatFichier='') and (TypeExp='') then
  BEGIN
  Readln(F,St) ;
  if Pos('¥',St)<>0 then St:=Copy(St,1,Pos('¥',St)-1) ;
  if (Copy(St,1,5)='00030') then BEGIN TypeExp:='FEC' ; FormatFichier:='EDI' ; END else
   if (Copy(St,1,5)='01030') then BEGIN TypeExp:='FBA' ; FormatFichier:='EDI' ; END else
    if (Copy(St,1,5)='02030') then
      BEGIN
      if Copy(St,34,3)='PCG' then TypeExp:='G' else
        if Copy(St,34,3)='PCA' then TypeExp:='A' else
          if Copy(St,34,3)='PCD' then TypeExp:='X' ;
      FormatFichier:='EDI' ;
      END else
  // formats exportés par saari,mp,rappro...
  // Comptes...
  if (((Copy(St,1,1)='X') or (Copy(St,1,1)='G') or (Copy(St,1,1)='A'))
    and ((Length(St)=69) or (Length(St)=146) or (Length(St)=145)) { 145 : Pour ORLI }
    or (Length(St)=189) or (Length(St)=212)) then
    BEGIN
    TypeExp:=Copy(St,1,1) ;
    Case Length(St) of
       69,146 : BEGIN TypeExp:=(Copy(St,1,1)) ; FormatFichier:='SAA' ; END ;
       189,212 : BEGIN TypeExp:='T' ; FormatFichier:='SAB' ; END ;
       END ;
    END else
    BEGIN
    // Ecritures, balances...
    Case Length(St) of
      112: BEGIN TypeExp:='FEC' ; FormatFichier:='MP' ; END ;
      132: BEGIN TypeExp:='FEC' ; FormatFichier:='RAP' ;END ;
      END ;

    While not Eof(f) do
      BEGIN
      if Pos('¥',St)<>0 then St:=Copy(St,1,Pos('¥',St)-1) ;
      if Copy(St,25,1)='E' then BEGIN TypeExp:='FEC' ; FormatFichier:='SN2' ; Break ; END
                           else if Copy(St,25,1)='X' then BEGIN TypeExp:='FEC' ; FormatFichier:='SAA' ; END ;
      Readln(F,St) ;
      END ;
    END ;
  END ;
CloseFile(F) ;
If Sn2Orli And (Copy(St,1,1)='X') Then
   BEGIN
   TypeExp:=(Copy(St,1,1)) ; FormatFichier:='SAA' ;
   END ;
if (FormatFichier<>'') and (TypeExp<>'') then
  if (Fmt<>FormatFichier) or (Trim(Lequel)<>TypeExp) then Exit ;
Result:=True ;
END ;


procedure TFFmtChoix.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFFmtChoix.BParamClick(Sender: TObject);
begin ParamImport (CurNat,Imp) ; end;

procedure TFFmtChoix.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if Parent is THPanel then Action:=caFree ;
end;

procedure TFFmtChoix.BScenarioClick(Sender: TObject);
begin
If Imp='X' Then ParamSupImport(IMP,CurNat,FFMt.Value,'',taModif,1)
           Else ParamSupExport(IMP,CurNat,FFMt.Value,'',taModif,1) ;
end;

procedure TFFmtChoix.FFmtChange(Sender: TObject);
begin
ScenarioVisible ;
end;

procedure TFFmtChoix.BFourCptClick(Sender: TObject);
begin
{$IFNDEF TOPGI}
ParamFourchetteImportCpt ;
{$ENDIF}
end;

end.
