unit FouCpt;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Hcompte, HRegCpte, Hctrls, ExtCtrls, ComCtrls, HTB97,
  Ent1, HEnt1, HPanel, UiUtil,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  hmsgbox,
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  Paramsoc, HSysMenu  ;

Procedure ParamFourchetteImportCpt ;

type
  TFFourImport = class(TForm)
    Pages: TPageControl;
    CptBil: TTabSheet;
    CptGes: TTabSheet;
    GbChar: TGroupBox;
    TSO_CHADEB1: THLabel;
    TSO_CHADEB2: THLabel;
    TSO_CHADEB3: THLabel;
    TSO_CHADEB4: THLabel;
    TSO_CHADEB5: THLabel;
    TSO_CHAFIN1: THLabel;
    TSO_CHAFIN2: THLabel;
    TSO_CHAFIN3: THLabel;
    TSO_CHAFIN4: THLabel;
    TSO_CHAFIN5: THLabel;
    SO_CHADEB1: THCpteEdit;
    SO_CHADEB2: THCpteEdit;
    SO_CHADEB3: THCpteEdit;
    SO_CHADEB4: THCpteEdit;
    SO_CHADEB5: THCpteEdit;
    SO_CHAFIN1: THCpteEdit;
    SO_CHAFIN2: THCpteEdit;
    SO_CHAFIN3: THCpteEdit;
    SO_CHAFIN4: THCpteEdit;
    SO_CHAFIN5: THCpteEdit;
    GbProd: TGroupBox;
    TSO_PRODEB1: THLabel;
    TSO_PRODEB2: THLabel;
    TSO_PRODEB3: THLabel;
    TSO_PRODEB4: THLabel;
    TSO_PRODEB5: THLabel;
    TSO_PROFIN1: THLabel;
    TSO_PROFIN2: THLabel;
    TSO_PROFIN3: THLabel;
    TSO_PROFIN4: THLabel;
    TSO_PROFIN5: THLabel;
    SO_PRODEB1: THCpteEdit;
    SO_PRODEB2: THCpteEdit;
    SO_PRODEB3: THCpteEdit;
    SO_PRODEB4: THCpteEdit;
    SO_PRODEB5: THCpteEdit;
    SO_PROFIN1: THCpteEdit;
    SO_PROFIN2: THCpteEdit;
    SO_PROFIN3: THCpteEdit;
    SO_PROFIN4: THCpteEdit;
    SO_PROFIN5: THCpteEdit;
    GroupBox4: TGroupBox;
    HLabel31: THLabel;
    HLabel32: THLabel;
    HLabel33: THLabel;
    HLabel34: THLabel;
    HLabel35: THLabel;
    HLabel36: THLabel;
    HLabel37: THLabel;
    HLabel38: THLabel;
    HLabel39: THLabel;
    HLabel40: THLabel;
    SO_CREDEB1: THCpteEdit;
    SO_CREDEB2: THCpteEdit;
    SO_CREDEB3: THCpteEdit;
    SO_CREDEB4: THCpteEdit;
    SO_CREDEB5: THCpteEdit;
    SO_CREFIN1: THCpteEdit;
    SO_CREFIN2: THCpteEdit;
    SO_CREFIN3: THCpteEdit;
    SO_CREFIN4: THCpteEdit;
    SO_CREFIN5: THCpteEdit;
    GroupBox3: TGroupBox;
    HLabel21: THLabel;
    HLabel22: THLabel;
    HLabel23: THLabel;
    HLabel24: THLabel;
    HLabel25: THLabel;
    HLabel26: THLabel;
    HLabel27: THLabel;
    HLabel28: THLabel;
    HLabel29: THLabel;
    HLabel30: THLabel;
    SO_DEBDEB1: THCpteEdit;
    SO_DEBDEB2: THCpteEdit;
    SO_DEBDEB3: THCpteEdit;
    SO_DEBDEB4: THCpteEdit;
    SO_DEBDEB5: THCpteEdit;
    SO_DEBFIN1: THCpteEdit;
    SO_DEBFIN2: THCpteEdit;
    SO_DEBFIN3: THCpteEdit;
    SO_DEBFIN4: THCpteEdit;
    SO_DEBFIN5: THCpteEdit;
    GroupBox1: TGroupBox;
    HLabel1: THLabel;
    HLabel2: THLabel;
    HLabel3: THLabel;
    HLabel6: THLabel;
    HLabel7: THLabel;
    HLabel8: THLabel;
    SO_BANDEB1: THCpteEdit;
    SO_BANDEB2: THCpteEdit;
    SO_BANDEB3: THCpteEdit;
    SO_BANFIN1: THCpteEdit;
    SO_BANFIN2: THCpteEdit;
    SO_BANFIN3: THCpteEdit;
    GroupBox2: TGroupBox;
    HLabel11: THLabel;
    HLabel12: THLabel;
    HLabel13: THLabel;
    HLabel14: THLabel;
    HLabel15: THLabel;
    HLabel16: THLabel;
    HLabel17: THLabel;
    HLabel18: THLabel;
    HLabel19: THLabel;
    HLabel20: THLabel;
    SO_IMMDEB1: THCpteEdit;
    SO_IMMDEB2: THCpteEdit;
    SO_IMMDEB3: THCpteEdit;
    SO_IMMDEB4: THCpteEdit;
    SO_IMMDEB5: THCpteEdit;
    SO_IMMFIN1: THCpteEdit;
    SO_IMMFIN2: THCpteEdit;
    SO_IMMFIN3: THCpteEdit;
    SO_IMMFIN4: THCpteEdit;
    SO_IMMFIN5: THCpteEdit;
    GroupBox5: TGroupBox;
    HLabel41: THLabel;
    HLabel46: THLabel;
    SO_CAIDEB1: THCpteEdit;
    SO_CAIFIN1: THCpteEdit;
    Dock: TDock97;
    HPB: TToolWindow97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    HM: THMsgBox;
    TabSheet1: TTabSheet;
    GroupBox6: TGroupBox;
    HLabel5: THLabel;
    HLabel9: THLabel;
    HLabel10: THLabel;
    HLabel42: THLabel;
    HLabel43: THLabel;
    HLabel44: THLabel;
    HLabel45: THLabel;
    HLabel47: THLabel;
    HLabel48: THLabel;
    HLabel49: THLabel;
    SO_CLIDEB1: THCpteEdit;
    SO_CLIDEB2: THCpteEdit;
    SO_CLIDEB3: THCpteEdit;
    SO_CLIDEB4: THCpteEdit;
    SO_CLIDEB5: THCpteEdit;
    SO_CLIFIN1: THCpteEdit;
    SO_CLIFIN2: THCpteEdit;
    SO_CLIFIN3: THCpteEdit;
    SO_CLIFIN4: THCpteEdit;
    SO_CLIFIN5: THCpteEdit;
    GroupBox7: TGroupBox;
    HLabel50: THLabel;
    HLabel51: THLabel;
    HLabel52: THLabel;
    HLabel53: THLabel;
    HLabel54: THLabel;
    HLabel55: THLabel;
    HLabel56: THLabel;
    HLabel57: THLabel;
    HLabel58: THLabel;
    HLabel59: THLabel;
    SO_FOUDEB1: THCpteEdit;
    SO_FOUDEB2: THCpteEdit;
    SO_FOUDEB3: THCpteEdit;
    SO_FOUDEB4: THCpteEdit;
    SO_FOUDEB5: THCpteEdit;
    SO_FOUFIN1: THCpteEdit;
    SO_FOUFIN2: THCpteEdit;
    SO_FOUFIN3: THCpteEdit;
    SO_FOUFIN4: THCpteEdit;
    SO_FOUFIN5: THCpteEdit;
    GroupBox12: TGroupBox;
    HLabel4: THLabel;
    HLabel60: THLabel;
    HLabel61: THLabel;
    HLabel62: THLabel;
    HLabel63: THLabel;
    HLabel64: THLabel;
    SO_SALDEB1: THCpteEdit;
    SO_SALDEB2: THCpteEdit;
    SO_SALDEB3: THCpteEdit;
    SO_SALFIN1: THCpteEdit;
    SO_SALFIN2: THCpteEdit;
    SO_SALFIN3: THCpteEdit;
    GroupBox13: TGroupBox;
    HLabel65: THLabel;
    HLabel66: THLabel;
    HLabel67: THLabel;
    HLabel68: THLabel;
    HLabel69: THLabel;
    HLabel70: THLabel;
    SO_DIVDEB1: THCpteEdit;
    SO_DIVDEB2: THCpteEdit;
    SO_DIVDEB3: THCpteEdit;
    SO_DIVFIN1: THCpteEdit;
    SO_DIVFIN2: THCpteEdit;
    SO_DIVFIN3: THCpteEdit;
    HMTrad: THSystemMenu;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BFermeClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure CptExit(Sender: TObject);
  private
    { Déclarations privées }
    Procedure ChargeFourchette ;
    Procedure SauveFourchette ;
    Function VerifFourchette : Boolean ;
    Procedure AlimFourchetteMemoire(Var Cha,Pro,DebD,DebC,Imm,Ban,Cai,Cli,Fou,Sal,Dvs : TFCha) ;
  public
    { Déclarations publiques }
  end;

implementation

uses ImpFicU,TimpFic ;

{$R *.DFM}

Procedure ParamFourchetteImportCpt ;
var FFR: TFFourImport;
    PP : THPanel ;
begin
FFR:=TFFourImport.Create(Application) ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FFR.ShowModal ;
    finally
     FFR.Free ;
    end ;
   END else
   BEGIN
   InitInside(FFR,PP) ;
   FFR.Show ;
   END ;
Screen.Cursor:=SyncrDefault ;
end ;

Function BourreLeDonc ( St : String ; LeType : TFichierBase ; Debut : Boolean) : string ;
var Lg,ll,i : Integer ;
    Bourre  : Char ;
begin
If LeType In [fbAxe1..fbAux,fbNatCpt] Then
   BEGIN
   Lg:=VH^.Cpta[LeType].Lg ;
   If LeType=fbNatCpt Then Bourre:=VH^.BourreLibre Else Bourre:=VH^.Cpta[LeType].Cb ;
   If Not Debut Then Bourre:='9' ;
   Result:=St ; ll:=Length(Result) ;
   If ll<Lg then for i:=ll+1 to Lg do Result:=Result+Bourre ;
   END Else Result:=St ;
end ;

Procedure TFFourImport.ChargeFourchette ;
Var
{$IFDEF SPEC302}
    Q : TQuery ;
{$ENDIF}
    DebD : TFCha ;
    DebC : TFCha ;
    Imm  : TFCha ;
    Ban  : TFcha ;
    Cai  : TFcha ;
    Cli  : TFcha ;
    Fou  : TFcha ;
    Sal  : TFcha ;
    Dvs  : TFcha ;
BEGIN
{$IFDEF SPEC302}
Q:=OpenSQL('SELECT * FROM SOCIETE',TRUE) ;
If Not Q.Eof Then
  BEGIN
  SO_CHADEB1.Text:=Q.FindField('SO_CHADEB1').AsString ; SO_CHAFIN1.Text:=Q.FindField('SO_CHAFIN1').AsString ;
  SO_CHADEB2.Text:=Q.FindField('SO_CHADEB2').AsString ; SO_CHAFIN2.Text:=Q.FindField('SO_CHAFIN2').AsString ;
  SO_CHADEB3.Text:=Q.FindField('SO_CHADEB3').AsString ; SO_CHAFIN3.Text:=Q.FindField('SO_CHAFIN3').AsString ;
  SO_CHADEB4.Text:=Q.FindField('SO_CHADEB4').AsString ; SO_CHAFIN4.Text:=Q.FindField('SO_CHAFIN4').AsString ;
  SO_CHADEB5.Text:=Q.FindField('SO_CHADEB5').AsString ; SO_CHAFIN5.Text:=Q.FindField('SO_CHAFIN5').AsString ;
  SO_PRODEB1.Text:=Q.FindField('SO_PRODEB1').AsString ; SO_PROFIN1.Text:=Q.FindField('SO_PROFIN1').AsString ;
  SO_PRODEB2.Text:=Q.FindField('SO_PRODEB2').AsString ; SO_PROFIN2.Text:=Q.FindField('SO_PROFIN2').AsString ;
  SO_PRODEB3.Text:=Q.FindField('SO_PRODEB3').AsString ; SO_PROFIN3.Text:=Q.FindField('SO_PROFIN3').AsString ;
  SO_PRODEB4.Text:=Q.FindField('SO_PRODEB4').AsString ; SO_PROFIN4.Text:=Q.FindField('SO_PROFIN4').AsString ;
  SO_PRODEB5.Text:=Q.FindField('SO_PRODEB5').AsString ; SO_PROFIN5.Text:=Q.FindField('SO_PROFIN5').AsString ;
  END ;
Ferme(Q) ;
{$ELSE}
SO_CHADEB1.Text:=GetParamSoc('SO_CHADEB1') ; SO_CHAFIN1.Text:=GetParamSoc('SO_CHAFIN1') ;
SO_CHADEB2.Text:=GetParamSoc('SO_CHADEB2') ; SO_CHAFIN2.Text:=GetParamSoc('SO_CHAFIN2') ;
SO_CHADEB3.Text:=GetParamSoc('SO_CHADEB3') ; SO_CHAFIN3.Text:=GetParamSoc('SO_CHAFIN3') ;
SO_CHADEB4.Text:=GetParamSoc('SO_CHADEB4') ; SO_CHAFIN4.Text:=GetParamSoc('SO_CHAFIN4') ;
SO_CHADEB5.Text:=GetParamSoc('SO_CHADEB5') ; SO_CHAFIN5.Text:=GetParamSoc('SO_CHAFIN5') ;
SO_PRODEB1.Text:=GetParamSoc('SO_PRODEB1') ; SO_PROFIN1.Text:=GetParamSoc('SO_PROFIN1') ;
SO_PRODEB2.Text:=GetParamSoc('SO_PRODEB2') ; SO_PROFIN2.Text:=GetParamSoc('SO_PROFIN2') ;
SO_PRODEB3.Text:=GetParamSoc('SO_PRODEB3') ; SO_PROFIN3.Text:=GetParamSoc('SO_PROFIN3') ;
SO_PRODEB4.Text:=GetParamSoc('SO_PRODEB4') ; SO_PROFIN4.Text:=GetParamSoc('SO_PROFIN4') ;
SO_PRODEB5.Text:=GetParamSoc('SO_PRODEB5') ; SO_PROFIN5.Text:=GetParamSoc('SO_PROFIN5') ;
{$ENDIF}
DechiffreFour(DebD,COTYPEIMP,'LD%',5) ;
SO_DEBDEB1.Text:=DebD[1].Deb ; SO_DEBFIN1.Text:=DebD[1].Fin ;
SO_DEBDEB2.Text:=DebD[2].Deb ; SO_DEBFIN2.Text:=DebD[2].Fin ;
SO_DEBDEB3.Text:=DebD[3].Deb ; SO_DEBFIN3.Text:=DebD[3].Fin ;
SO_DEBDEB4.Text:=DebD[4].Deb ; SO_DEBFIN4.Text:=DebD[4].Fin ;
SO_DEBDEB5.Text:=DebD[5].Deb ; SO_DEBFIN5.Text:=DebD[5].Fin ;
DechiffreFour(DebC,COTYPEIMP,'LC%',5) ;
SO_CREDEB1.Text:=DebC[1].Deb ; SO_CREFIN1.Text:=DebC[1].Fin ;
SO_CREDEB2.Text:=DebC[2].Deb ; SO_CREFIN2.Text:=DebC[2].Fin ;
SO_CREDEB3.Text:=DebC[3].Deb ; SO_CREFIN3.Text:=DebC[3].Fin ;
SO_CREDEB4.Text:=DebC[4].Deb ; SO_CREFIN4.Text:=DebC[4].Fin ;
SO_CREDEB5.Text:=DebC[5].Deb ; SO_CREFIN5.Text:=DebC[5].Fin ;
DechiffreFour(Imm,COTYPEIMP,'IM%',5) ;
SO_IMMDEB1.Text:=Imm[1].Deb ; SO_IMMFIN1.Text:=Imm[1].Fin ;
SO_IMMDEB2.Text:=Imm[2].Deb ; SO_IMMFIN2.Text:=Imm[2].Fin ;
SO_IMMDEB3.Text:=Imm[3].Deb ; SO_IMMFIN3.Text:=Imm[3].Fin ;
SO_IMMDEB4.Text:=Imm[4].Deb ; SO_IMMFIN4.Text:=Imm[4].Fin ;
SO_IMMDEB5.Text:=Imm[5].Deb ; SO_IMMFIN5.Text:=Imm[5].Fin ;
DechiffreFour(Ban,COTYPEIMP,'BA%',3) ;
SO_BANDEB1.Text:=Ban[1].Deb ; SO_BANFIN1.Text:=Ban[1].Fin ;
SO_BANDEB2.Text:=Ban[2].Deb ; SO_BANFIN2.Text:=Ban[2].Fin ;
SO_BANDEB3.Text:=Ban[3].Deb ; SO_BANFIN3.Text:=Ban[3].Fin ;
DechiffreFour(Cai,COTYPEIMP,'CA%',1) ;
SO_CAIDEB1.Text:=Cai[1].Deb ; SO_CAIFIN1.Text:=Cai[1].Fin ;
DechiffreFour(Cli,COTYPEIMP,'CL%',5) ;
SO_CLIDEB1.Text:=Cli[1].Deb ; SO_CLIFIN1.Text:=Cli[1].Fin ;
SO_CLIDEB2.Text:=Cli[2].Deb ; SO_CLIFIN2.Text:=Cli[2].Fin ;
SO_CLIDEB3.Text:=Cli[3].Deb ; SO_CLIFIN3.Text:=Cli[3].Fin ;
SO_CLIDEB4.Text:=Cli[4].Deb ; SO_CLIFIN4.Text:=Cli[4].Fin ;
SO_CLIDEB5.Text:=Cli[5].Deb ; SO_CLIFIN5.Text:=Cli[5].Fin ;
DechiffreFour(Fou,COTYPEIMP,'FO%',5) ;
SO_FOUDEB1.Text:=Fou[1].Deb ; SO_FOUFIN1.Text:=Fou[1].Fin ;
SO_FOUDEB2.Text:=Fou[2].Deb ; SO_FOUFIN2.Text:=Fou[2].Fin ;
SO_FOUDEB3.Text:=Fou[3].Deb ; SO_FOUFIN3.Text:=Fou[3].Fin ;
SO_FOUDEB4.Text:=Fou[4].Deb ; SO_FOUFIN4.Text:=Fou[4].Fin ;
SO_FOUDEB5.Text:=Fou[5].Deb ; SO_FOUFIN5.Text:=Fou[5].Fin ;
DechiffreFour(Sal,COTYPEIMP,'SA%',3) ;
SO_SALDEB1.Text:=Sal[1].Deb ; SO_SALFIN1.Text:=Sal[1].Fin ;
SO_SALDEB2.Text:=Sal[2].Deb ; SO_SALFIN2.Text:=Sal[2].Fin ;
SO_SALDEB3.Text:=Sal[3].Deb ; SO_SALFIN3.Text:=Sal[3].Fin ;
DechiffreFour(Dvs,COTYPEIMP,'DI%',3) ;
SO_DIVDEB1.Text:=Dvs[1].Deb ; SO_DIVFIN1.Text:=Dvs[1].Fin ;
SO_DIVDEB2.Text:=Dvs[2].Deb ; SO_DIVFIN2.Text:=Dvs[2].Fin ;
SO_DIVDEB3.Text:=Dvs[3].Deb ; SO_DIVFIN3.Text:=Dvs[3].Fin ;
END ;

Procedure TFFourImport.AlimFourchetteMemoire(Var Cha,Pro,DebD,DebC,Imm,Ban,Cai,Cli,Fou,Sal,Dvs : TFCha) ;
BEGIN
Fillchar(Cha,SizeOf(Cha),#0) ; Fillchar(Pro,SizeOf(Pro),#0) ;
Fillchar(DebD,SizeOf(DebD),#0) ; Fillchar(DebC,SizeOf(DebC),#0) ;
Fillchar(Imm,SizeOf(Imm),#0) ; Fillchar(Ban,SizeOf(Ban),#0) ;
Fillchar(Cai,SizeOf(Cai),#0) ;
Fillchar(Sal,SizeOf(Sal),#0) ; Fillchar(Dvs,SizeOf(Dvs),#0) ;
Cha[1].Deb:=Trim(SO_CHADEB1.Text) ; Cha[1].Fin:=Trim(SO_CHAFIN1.Text) ;
Cha[2].Deb:=Trim(SO_CHADEB2.Text) ; Cha[2].Fin:=Trim(SO_CHAFIN2.Text) ;
Cha[3].Deb:=Trim(SO_CHADEB3.Text) ; Cha[3].Fin:=Trim(SO_CHAFIN3.Text) ;
Cha[4].Deb:=Trim(SO_CHADEB4.Text) ; Cha[4].Fin:=Trim(SO_CHAFIN4.Text) ;
Cha[5].Deb:=Trim(SO_CHADEB5.Text) ; Cha[5].Fin:=Trim(SO_CHAFIN5.Text) ;
Pro[1].Deb:=Trim(SO_PRODEB1.Text) ; Pro[1].Fin:=Trim(SO_PROFIN1.Text) ;
Pro[2].Deb:=Trim(SO_PRODEB2.Text) ; Pro[2].Fin:=Trim(SO_PROFIN2.Text) ;
Pro[3].Deb:=Trim(SO_PRODEB3.Text) ; Pro[3].Fin:=Trim(SO_PROFIN3.Text) ;
Pro[4].Deb:=Trim(SO_PRODEB4.Text) ; Pro[4].Fin:=Trim(SO_PROFIN4.Text) ;
Pro[5].Deb:=Trim(SO_PRODEB5.Text) ; Pro[5].Fin:=Trim(SO_PRODEB5.Text) ;
DebD[1].Deb:=Trim(SO_DEBDEB1.text) ; DebD[1].Fin:=Trim(SO_DEBFIN1.text) ;
DebD[2].Deb:=Trim(SO_DEBDEB2.text) ; DebD[2].Fin:=Trim(SO_DEBFIN2.text) ;
DebD[3].Deb:=Trim(SO_DEBDEB3.text) ; DebD[3].Fin:=Trim(SO_DEBFIN3.text) ;
DebD[4].Deb:=Trim(SO_DEBDEB4.text) ; DebD[4].Fin:=Trim(SO_DEBFIN4.text) ;
DebD[5].Deb:=Trim(SO_DEBDEB5.text) ; DebD[5].Fin:=Trim(SO_DEBFIN5.text) ;
DebC[1].Deb:=Trim(SO_CREDEB1.text) ; DebC[1].Fin:=Trim(SO_CREFIN1.text) ;
DebC[2].Deb:=Trim(SO_CREDEB2.text) ; DebC[2].Fin:=Trim(SO_CREFIN2.text) ;
DebC[3].Deb:=Trim(SO_CREDEB3.text) ; DebC[3].Fin:=Trim(SO_CREFIN3.text) ;
DebC[4].Deb:=Trim(SO_CREDEB4.text) ; DebC[4].Fin:=Trim(SO_CREFIN4.text) ;
DebC[5].Deb:=Trim(SO_CREDEB5.text) ; DebC[5].Fin:=Trim(SO_CREFIN5.text) ;
Imm[1].Deb:=Trim(SO_IMMDEB1.text) ; Imm[1].Fin:=Trim(SO_IMMFIN1.text) ;
Imm[2].Deb:=Trim(SO_IMMDEB2.text) ; Imm[2].Fin:=Trim(SO_IMMFIN2.text) ;
Imm[3].Deb:=Trim(SO_IMMDEB3.text) ; Imm[3].Fin:=Trim(SO_IMMFIN3.text) ;
Imm[4].Deb:=Trim(SO_IMMDEB4.text) ; Imm[4].Fin:=Trim(SO_IMMFIN4.text) ;
Imm[5].Deb:=Trim(SO_IMMDEB5.text) ; Imm[5].Fin:=Trim(SO_IMMFIN5.text) ;
Ban[1].Deb:=Trim(SO_BANDEB1.text) ; Ban[1].Fin:=Trim(SO_BANFIN1.text) ;
Ban[2].Deb:=Trim(SO_BANDEB2.text) ; Ban[2].Fin:=Trim(SO_BANFIN2.text) ;
Ban[3].Deb:=Trim(SO_BANDEB3.text) ; Ban[3].Fin:=Trim(SO_BANFIN3.text) ;
Cai[1].Deb:=Trim(SO_CAIDEB1.text) ; Cai[1].Fin:=Trim(SO_CAIFIN1.text) ;
Cli[1].Deb:=Trim(SO_CLIDEB1.text) ; Cli[1].Fin:=Trim(SO_CLIFIN1.text) ;
Cli[2].Deb:=Trim(SO_CLIDEB2.text) ; Cli[2].Fin:=Trim(SO_CLIFIN2.text) ;
Cli[3].Deb:=Trim(SO_CLIDEB3.text) ; Cli[3].Fin:=Trim(SO_CLIFIN3.text) ;
Cli[4].Deb:=Trim(SO_CLIDEB4.text) ; Cli[4].Fin:=Trim(SO_CLIFIN4.text) ;
Cli[5].Deb:=Trim(SO_CLIDEB5.text) ; Cli[5].Fin:=Trim(SO_CLIFIN5.text) ;
Fou[1].Deb:=Trim(SO_FOUDEB1.text) ; Fou[1].Fin:=Trim(SO_FOUFIN1.text) ;
Fou[2].Deb:=Trim(SO_FOUDEB2.text) ; Fou[2].Fin:=Trim(SO_FOUFIN2.text) ;
Fou[3].Deb:=Trim(SO_FOUDEB3.text) ; Fou[3].Fin:=Trim(SO_FOUFIN3.text) ;
Fou[4].Deb:=Trim(SO_FOUDEB4.text) ; Fou[4].Fin:=Trim(SO_FOUFIN4.text) ;
Fou[5].Deb:=Trim(SO_FOUDEB5.text) ; Fou[5].Fin:=Trim(SO_FOUFIN5.text) ;
Sal[1].Deb:=Trim(SO_SALDEB1.text) ; Sal[1].Fin:=Trim(SO_SALFIN1.text) ;
Sal[2].Deb:=Trim(SO_SALDEB2.text) ; Sal[2].Fin:=Trim(SO_SALFIN2.text) ;
Sal[3].Deb:=Trim(SO_SALDEB3.text) ; Sal[3].Fin:=Trim(SO_SALFIN3.text) ;
Dvs[1].Deb:=Trim(SO_DIVDEB1.text) ; Dvs[1].Fin:=Trim(SO_DIVFIN1.text) ;
Dvs[2].Deb:=Trim(SO_DIVDEB2.text) ; Dvs[2].Fin:=Trim(SO_DIVFIN2.text) ;
Dvs[3].Deb:=Trim(SO_DIVDEB3.text) ; Dvs[3].Fin:=Trim(SO_DIVFIN3.text) ;
END ;

Procedure InsertUneFouchette(Q : TQuery ; St : String ; MaxE : Integer ; F : TFCha) ;
Var i : Integer ;
BEGIN
For i:=1 To MaxE Do If (F[i].Deb<>'') And (F[i].Fin<>'' ) Then
  BEGIN
  Q.Insert ;
  InitNew(Q) ;
  Q.FindField('CC_TYPE').AsString:=COTYPEIMP ;
  Q.FindField('CC_CODE').AsString:=St+IntToStr(i) ;
  Q.FindField('CC_LIBELLE').AsString:=F[i].Deb ;
  Q.FindField('CC_ABREGE').AsString:=F[i].Fin ;
  Q.Post ;
  END ;
END ;

Procedure TFFourImport.SauveFourchette ;
Var Q : TQuery ;
    Cha  : TFCha ;
    Pro  : TFCha ;
    DebD : TFCha ;
    DebC : TFCha ;
    Imm  : TFCha ;
    Ban  : TFcha ;
    Cai  : TFcha ;
    Cli  : TFcha ;
    Fou  : TFcha ;
    Sal  : TFcha ;
    Dvs  : TFcha ;
BEGIN
{$IFDEF SPEC302}
Q:=OpenSQL('SELECT * FROM SOCIETE',FALSE) ;
If Not Q.Eof Then
  BEGIN
  Q.Edit ;
  Q.FindField('SO_CHADEB1').AsString:=SO_CHADEB1.Text ; Q.FindField('SO_CHAFIN1').AsString:=SO_CHAFIN1.Text ;
  Q.FindField('SO_CHADEB2').AsString:=SO_CHADEB2.Text ; Q.FindField('SO_CHAFIN2').AsString:=SO_CHAFIN2.Text ;
  Q.FindField('SO_CHADEB3').AsString:=SO_CHADEB3.Text ; Q.FindField('SO_CHAFIN3').AsString:=SO_CHAFIN3.Text ;
  Q.FindField('SO_CHADEB4').AsString:=SO_CHADEB4.Text ; Q.FindField('SO_CHAFIN4').AsString:=SO_CHAFIN4.Text ;
  Q.FindField('SO_CHADEB5').AsString:=SO_CHADEB5.Text ; Q.FindField('SO_CHAFIN5').AsString:=SO_CHAFIN5.Text ;
  Q.FindField('SO_PRODEB1').AsString:=SO_PRODEB1.Text ; Q.FindField('SO_PROFIN1').AsString:=SO_PROFIN1.Text ;
  Q.FindField('SO_PRODEB2').AsString:=SO_PRODEB2.Text ; Q.FindField('SO_PROFIN2').AsString:=SO_PROFIN2.Text ;
  Q.FindField('SO_PRODEB3').AsString:=SO_PRODEB3.Text ; Q.FindField('SO_PROFIN3').AsString:=SO_PROFIN3.Text ;
  Q.FindField('SO_PRODEB4').AsString:=SO_PRODEB4.Text ; Q.FindField('SO_PROFIN4').AsString:=SO_PROFIN4.Text ;
  Q.FindField('SO_PRODEB5').AsString:=SO_PRODEB5.Text ; Q.FindField('SO_PROFIN5').AsString:=SO_PRODEB5.Text ;
  Q.Post ;
  END ;
Ferme(Q) ;
{$ELSE}
SetParamSoc('SO_CHADEB1',SO_CHADEB1.Text) ; SetParamSoc('SO_CHAFIN1',SO_CHAFIN1.Text) ;
SetParamSoc('SO_CHADEB2',SO_CHADEB2.Text) ; SetParamSoc('SO_CHAFIN2',SO_CHAFIN2.Text) ;
SetParamSoc('SO_CHADEB3',SO_CHADEB3.Text) ; SetParamSoc('SO_CHAFIN3',SO_CHAFIN3.Text) ;
SetParamSoc('SO_CHADEB4',SO_CHADEB4.Text) ; SetParamSoc('SO_CHAFIN4',SO_CHAFIN4.Text) ;
SetParamSoc('SO_CHADEB5',SO_CHADEB5.Text) ; SetParamSoc('SO_CHAFIN5',SO_CHAFIN5.Text) ;
SetParamSoc('SO_PRODEB1',SO_PRODEB1.Text) ; SetParamSoc('SO_PROFIN1',SO_PROFIN1.Text) ;
SetParamSoc('SO_PRODEB2',SO_PRODEB2.Text) ; SetParamSoc('SO_PROFIN2',SO_PROFIN2.Text) ;
SetParamSoc('SO_PRODEB3',SO_PRODEB3.Text) ; SetParamSoc('SO_PROFIN3',SO_PROFIN3.Text) ;
SetParamSoc('SO_PRODEB4',SO_PRODEB4.Text) ; SetParamSoc('SO_PROFIN4',SO_PROFIN4.Text) ;
SetParamSoc('SO_PRODEB5',SO_PRODEB5.Text) ; SetParamSoc('SO_PROFIN5',SO_PRODEB5.Text) ;
{$ENDIF}
ExecuteSQL('DELETE FROM CHOIXCOD WHERE CC_TYPE="'+COTYPEIMP+'" AND CC_CODE<>"NOM"') ;
AlimFourchetteMemoire(Cha,Pro,DebD,DebC,Imm,Ban,Cai,Cli,Fou,Sal,Dvs) ;
Q:=OpenSQL('Select * from CHOIXCOD WHERE CC_TYPE="'+COTYPEIMP+'"',FALSE) ;
InsertUneFouchette(Q,'LD',5,DebD) ;
InsertUneFouchette(Q,'LC',5,DebC) ;
InsertUneFouchette(Q,'IM',5,Imm) ;
InsertUneFouchette(Q,'BA',3,Ban) ;
InsertUneFouchette(Q,'CA',1,Cai) ;
InsertUneFouchette(Q,'CL',5,Cli) ;
InsertUneFouchette(Q,'FO',5,Fou) ;
InsertUneFouchette(Q,'SA',3,Sal) ;
InsertUneFouchette(Q,'DI',3,Dvs) ;
Ferme(Q) ;
END ;

Function FourchetteOk(F : TFCha ; MaxE : Integer) : Boolean ;
Var i : Integer ;
BEGIN
Result:=TRUE ;
For i:=1 To MaxE Do If Trim(F[i].Deb)>Trim(F[i].Fin) Then BEGIN Result:=FALSE ; Exit ; END ;
END ;

Function Intersection(Deb,Fin : String ; F2 : TFCha ; Max2 : Integer) : Boolean ;
Var i : Integer ;
BEGIN
Result:=FALSE ;
For i:=1 To Max2 Do
  If ((Deb<>'') And (Deb>=F2[i].Deb) And (Deb<=F2[i].Fin)) Or
     ((Fin<>'') And (Fin>=F2[i].Deb) And (Fin<=F2[i].Fin)) Or
     ((F2[i].Deb<>'') And (F2[i].Deb>=Deb) And (F2[i].Deb<=Fin)) Or
     ((F2[i].Fin<>'') And (F2[i].Fin>=Deb) And (F2[i].Fin<=Fin)) Then BEGIN Result:=TRUE ; Exit ; END ;
END ;

Function FourchetteCoherente(F1,F2 : TFCha ; Max1,Max2 : Integer) : Boolean ;
Var i : Integer ;
BEGIN
Result:=TRUE ;
For i:=1 To Max1 Do
  BEGIN
  If Intersection(F1[i].Deb,F1[i].Fin,F2,Max2) Then BEGIN Result:=FALSE ; Exit ; END ;
  END ;
END ;

Function TFFourImport.VerifFourchette : Boolean ;
Var Cha  : TFCha ;
    Pro  : TFCha ;
    DebD : TFCha ;
    DebC : TFCha ;
    Imm  : TFCha ;
    Ban  : TFcha ;
    Cai  : TFcha ;
    Cli  : TFcha ;
    Fou  : TFcha ;
    Sal  : TFcha ;
    Dvs  : TFcha ;
    i : Integer ;
    St : String ;
BEGIN
Result:=FALSE ;
AlimFourchetteMemoire(Cha,Pro,DebD,DebC,Imm,Ban,Cai,Cli,Fou,Sal,Dvs) ;
i:=0 ;
If Not FourchetteOk(Cha,5) Then i:=1 ;
If (i=0) And (Not FourchetteOk(Pro,5)) Then i:=2 ;
If (i=0) And (Not FourchetteOk(DebD,5)) Then i:=3 ;
If (i=0) And (Not FourchetteOk(DebC,5)) Then i:=4 ;
If (i=0) And (Not FourchetteOk(Imm,5)) Then i:=5 ;
If (i=0) And (Not FourchetteOk(Ban,3)) Then i:=6 ;
If (i=0) And (Not FourchetteOk(Cai,1)) Then i:=7 ;
If (i=0) And (Not FourchetteOk(Cli,5)) Then i:=8 ;
If (i=0) And (Not FourchetteOk(Fou,5)) Then i:=9 ;
If (i=0) And (Not FourchetteOk(Sal,3)) Then i:=10 ;
If (i=0) And (Not FourchetteOk(Dvs,3)) Then i:=11 ;
If (i=0) And (Not FourchetteCoherente(Cha,Pro,5,5)) Then BEGIN i:=100 ; St:='Charges / Produits' ; END ;
If (i=0) And (Not FourchetteCoherente(Cha,DebD,5,5)) Then BEGIN i:=100 ; St:='Charges / Lettrables (débiteurs)' ; END ;
If (i=0) And (Not FourchetteCoherente(Cha,DebC,5,5)) Then BEGIN i:=100 ; St:='Charges / Lettrables (créditeurs)' ; END ;
If (i=0) And (Not FourchetteCoherente(Cha,Imm,5,5)) Then BEGIN i:=100 ; St:='Charges / Immobilisations' ; END ;
If (i=0) And (Not FourchetteCoherente(Cha,Ban,5,3)) Then BEGIN i:=100 ; St:='Charges / Banques' ; END ;
If (i=0) And (Not FourchetteCoherente(Cha,Cai,5,1)) Then BEGIN i:=100 ; St:='Charges / Caisses' ; END ;
If (i=0) And (Not FourchetteCoherente(Cha,Cli,5,5)) Then BEGIN i:=100 ; St:='Charges / Collectifs clients' ; END ;
If (i=0) And (Not FourchetteCoherente(Cha,Fou,5,5)) Then BEGIN i:=100 ; St:='Charges / Collectifs fournisseurs' ; END ;
If (i=0) And (Not FourchetteCoherente(Cha,Sal,5,5)) Then BEGIN i:=100 ; St:='Charges / Collectifs salariés' ; END ;
If (i=0) And (Not FourchetteCoherente(Cha,Dvs,5,3)) Then BEGIN i:=100 ; St:='Charges / Collectifs divers' ; END ;
If (i=0) And (Not FourchetteCoherente(Pro,DebD,5,5)) Then BEGIN i:=100 ; St:='Produits / Lettrables (débiteurs)' ; END ;
If (i=0) And (Not FourchetteCoherente(Pro,DebC,5,5)) Then BEGIN i:=100 ; St:='Produits / Lettrables (créditeurs)' ; END ;
If (i=0) And (Not FourchetteCoherente(Pro,Imm,5,5)) Then BEGIN i:=100 ; St:='Produits / Immobilisations' ; END ;
If (i=0) And (Not FourchetteCoherente(Pro,Ban,5,3)) Then BEGIN i:=100 ; St:='Produits / Banques' ; END ;
If (i=0) And (Not FourchetteCoherente(Pro,Cai,5,1)) Then BEGIN i:=100 ; St:='Produits / Caisses' ; END ;
If (i=0) And (Not FourchetteCoherente(Pro,Cli,5,5)) Then BEGIN i:=100 ; St:='Produits / Collectifs clients' ; END ;
If (i=0) And (Not FourchetteCoherente(Pro,Fou,5,5)) Then BEGIN i:=100 ; St:='Produits / Collectifs fournisseurs' ; END ;
If (i=0) And (Not FourchetteCoherente(Pro,Sal,5,5)) Then BEGIN i:=100 ; St:='Produits / Collectifs salariés' ; END ;
If (i=0) And (Not FourchetteCoherente(Pro,Dvs,5,3)) Then BEGIN i:=100 ; St:='Produits / Collectifs divers' ; END ;
If (i=0) And (Not FourchetteCoherente(DebD,DebC,5,5)) Then BEGIN i:=100 ; St:='Lettrables (débiteurs) / Lettrables (créditeurs)' ; END ;
If (i=0) And (Not FourchetteCoherente(DebD,Imm,5,5)) Then BEGIN i:=100 ; St:='Lettrables (débiteurs) / Immobilisations' ; END ;
If (i=0) And (Not FourchetteCoherente(DebD,Ban,5,3)) Then BEGIN i:=100 ; St:='Lettrables (débiteurs) / Banques' ; END ;
If (i=0) And (Not FourchetteCoherente(DebD,Cai,5,1)) Then BEGIN i:=100 ; St:='Lettrables (débiteurs) / Caisses' ; END ;
If (i=0) And (Not FourchetteCoherente(DebD,Cli,5,5)) Then BEGIN i:=100 ; St:='Lettrables (débiteurs) / Collectifs clients' ; END ;
If (i=0) And (Not FourchetteCoherente(DebD,Fou,5,5)) Then BEGIN i:=100 ; St:='Lettrables (débiteurs) / Collectifs fournisseurs' ; END ;
If (i=0) And (Not FourchetteCoherente(DebD,Sal,5,5)) Then BEGIN i:=100 ; St:='Lettrables (débiteurs) / Collectifs salariés' ; END ;
If (i=0) And (Not FourchetteCoherente(DebD,Dvs,5,3)) Then BEGIN i:=100 ; St:='Lettrables (débiteurs) / Collectifs divers' ; END ;
If (i=0) And (Not FourchetteCoherente(DebC,Imm,5,5)) Then BEGIN i:=100 ; St:='Lettrables (créditeurs) / Immobilisations' ; END ;
If (i=0) And (Not FourchetteCoherente(DebC,Ban,5,3)) Then BEGIN i:=100 ; St:='Lettrables (créditeurs) / Banques' ; END ;
If (i=0) And (Not FourchetteCoherente(DebC,Cai,5,1)) Then BEGIN i:=100 ; St:='Lettrables (créditeurs) / Caisses' ; END ;
If (i=0) And (Not FourchetteCoherente(DebC,Cli,5,5)) Then BEGIN i:=100 ; St:='Lettrables (créditeurs) / Collectifs clients' ; END ;
If (i=0) And (Not FourchetteCoherente(DebC,Fou,5,5)) Then BEGIN i:=100 ; St:='Lettrables (créditeurs) / Collectifs fournisseurs' ; END ;
If (i=0) And (Not FourchetteCoherente(DebC,Sal,5,5)) Then BEGIN i:=100 ; St:='Lettrables (créditeurs) / Collectifs salariés' ; END ;
If (i=0) And (Not FourchetteCoherente(DebC,Dvs,5,3)) Then BEGIN i:=100 ; St:='Lettrables (créditeurs) / Collectifs divers' ; END ;
If (i=0) And (Not FourchetteCoherente(Imm,Ban,5,3)) Then BEGIN i:=100 ; St:='Immobilisations / Banques' ; END ;
If (i=0) And (Not FourchetteCoherente(Imm,Cai,5,1)) Then BEGIN i:=100 ; St:='Immobilisations / Caisses' ; END ;
If (i=0) And (Not FourchetteCoherente(Imm,Cli,5,5)) Then BEGIN i:=100 ; St:='Immobilisations / Collectifs clients' ; END ;
If (i=0) And (Not FourchetteCoherente(Imm,Fou,5,5)) Then BEGIN i:=100 ; St:='Immobilisations / Collectifs fournisseurs' ; END ;
If (i=0) And (Not FourchetteCoherente(Imm,Sal,5,5)) Then BEGIN i:=100 ; St:='Immobilisations / Collectifs salariés' ; END ;
If (i=0) And (Not FourchetteCoherente(Imm,Dvs,5,3)) Then BEGIN i:=100 ; St:='Immobilisations / Collectifs divers' ; END ;
If (i=0) And (Not FourchetteCoherente(Ban,Cai,3,1)) Then BEGIN i:=100 ; St:='Banques / Caisses' ; END ;
If (i=0) And (Not FourchetteCoherente(Ban,Cli,3,5)) Then BEGIN i:=100 ; St:='Banques / Collectifs clients' ; END ;
If (i=0) And (Not FourchetteCoherente(Ban,Fou,3,5)) Then BEGIN i:=100 ; St:='Banques / Collectifs fournisseurs' ; END ;
If (i=0) And (Not FourchetteCoherente(Ban,Sal,3,5)) Then BEGIN i:=100 ; St:='Banques / Collectifs salariés' ; END ;
If (i=0) And (Not FourchetteCoherente(Ban,Dvs,3,3)) Then BEGIN i:=100 ; St:='Banques / Collectifs divers' ; END ;
If (i=0) And (Not FourchetteCoherente(Cai,Cli,1,5)) Then BEGIN i:=100 ; St:='Caisses / Collectifs clients' ; END ;
If (i=0) And (Not FourchetteCoherente(Cai,Fou,1,5)) Then BEGIN i:=100 ; St:='Caisses / Collectifs fournisseurs' ; END ;
If (i=0) And (Not FourchetteCoherente(Cai,Sal,1,5)) Then BEGIN i:=100 ; St:='Caisses / Collectifs salariés' ; END ;
If (i=0) And (Not FourchetteCoherente(Cai,Dvs,1,3)) Then BEGIN i:=100 ; St:='Caisses / Collectifs divers' ; END ;
If (i=0) And (Not FourchetteCoherente(Cli,Fou,5,5)) Then BEGIN i:=100 ; St:='Collectifs clients / Collectifs fournisseurs' ; END ;
If (i=0) And (Not FourchetteCoherente(Cli,Sal,5,5)) Then BEGIN i:=100 ; St:='Collectifs clients / Collectifs salariés' ; END ;
If (i=0) And (Not FourchetteCoherente(Cli,Dvs,5,3)) Then BEGIN i:=100 ; St:='Collectifs clients / Collectifs divers' ; END ;
If (i=0) And (Not FourchetteCoherente(Fou,Sal,5,5)) Then BEGIN i:=100 ; St:='Collectifs fournisseurs / Collectifs salariés' ; END ;
If (i=0) And (Not FourchetteCoherente(Fou,Dvs,5,3)) Then BEGIN i:=100 ; St:='Collectifs fournisseurs / Collectifs divers' ; END ;
If (i=0) And (Not FourchetteCoherente(Sal,Dvs,3,3)) Then BEGIN i:=100 ; St:='Collectifs salariés / Collectifs divers' ; END ;
If i<>0 Then
  BEGIN
  If i=100 Then HShowMessage('0;Fourchettes de comptes;Fourchettes en intersection ('+St+');E;O;O;O;','','')
           Else HM.Execute(i,'','') ;
  END Else Result:=TRUE ;
END ;

procedure TFFourImport.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if Parent is THPanel then Action:=caFree ;
end;

procedure TFFourImport.BFermeClick(Sender: TObject);
begin
Close ;
end;

procedure TFFourImport.BValiderClick(Sender: TObject);
begin
If HM.Execute(0,'','')=mrYes Then
  BEGIN
  If VerifFourchette Then BEGIN SauveFourchette ; Close ; END ;
  END ;
end;

procedure TFFourImport.FormShow(Sender: TObject);
begin
Pages.ActivePage:=Pages.Pages[0] ;
ChargeFourchette ;
end;

procedure TFFourImport.CptExit(Sender: TObject);
Var St : String ;
    Deb : Boolean ;                   
begin
St:=THCpteEdit(Sender).Text ; If Trim(St)='' Then Exit ;
Deb:=Pos('FIN',THCpteEdit(Sender).Name)=0 ;
St:=BourreLeDonc(St,fbGene,Deb) ;
THCpteEdit(Sender).Text:=St ;
end;

end.
