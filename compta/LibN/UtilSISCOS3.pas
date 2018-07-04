unit UtilSISCOS3;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Hctrls, ComCtrls, HTB97, HSysMenu, UiUtil, HPanel, Ent1, Hent1,
  hmsgbox, FileCtrl, UTOB,
 {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  Mask, DBCtrls, Spin, HDB, ExtCtrls ;

Procedure RecupSISCOS3(Recup : Boolean) ;

type
  TFParamRecupSISCO = class(TForm)
    Dock971: TDock97;
    HPB: TToolWindow97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    Pages: TPageControl;
    PStandards: TTabSheet;
    Pcomptes: TTabSheet;
    PDivers: TTabSheet;
    GP1: TGroupBox;
    Label1: TLabel;
    FMasquePrefTRT: TEdit;
    Label3: TLabel;
    Label2: TLabel;
    FRepTRT: TEdit;
    FMasqueSuffTRT: TEdit;
    BDirTRT: TButton;
    Msg: THMsgBox;
    HMTrad: THSystemMenu;
    GP2: TGroupBox;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    FMasquePrefCLI: TEdit;
    FRepCli: TEdit;
    FMasqueSuffCli: TEdit;
    BDirCli: TButton;
    GbChar: TGroupBox;
    CBAnaCha: TCheckBox;
    GbProd: TGroupBox;
    CBAnaPro: TCheckBox;
    FOk1: TCheckBox;
    FOk2: TCheckBox;
    FBigVol: TCheckBox;
    TFNomChamp: TLabel;
    FTypeLettrage: THValComboBox;
    PSupp: TTabSheet;
    GP4: TGroupBox;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    FMasquePrefImmo: TEdit;
    FRepImmo: TEdit;
    FMasqueSuffImmo: TEdit;
    BDirImmo: TButton;
    FOk4: TCheckBox;
    GP5: TGroupBox;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    FMasqueprefSit: TEdit;
    FRepSit: TEdit;
    FMasqueSuffSit: TEdit;
    BDirSit: TButton;
    FOk5: TCheckBox;
    GP6: TGroupBox;
    TSO_LGCPTEGEN: THLabel;
    TSO_LGCPTEAUX: THLabel;
    TSO_BOURRETIERS: THLabel;
    TSO_BOURREGEN: THLabel;
    LGGen: TSpinEdit;
    LGAux: TSpinEdit;
    BOURREAUX: TEdit;
    BOURREGEN: TEdit;
    TSO_LGCPTESECT: THLabel;
    LgAna: TSpinEdit;
    TSO_BOURRESECT: THLabel;
    BOURREANA: TEdit;
    FOk6: TCheckBox;
    Sauve: TSaveDialog;
    BRecupSISCO: TToolbarButton97;
    FStat: THValComboBox;
    Label16: TLabel;
    GroupBox1: TGroupBox;
    TFCodeLettre2: THLabel;
    HLabel1: THLabel;
    HLabel2: THLabel;
    HLabel3: THLabel;
    HLabel4: THLabel;
    HLabel5: THLabel;
    BZoomJalAch: TToolbarButton97;
    BZoomJalVen: TToolbarButton97;
    BZoomCptLett: TToolbarButton97;
    BZoomCptPoint: TToolbarButton97;
    FCptEcart: TEdit;
    FJalAch: TEdit;
    FJalVen: TEdit;
    FJalAN: TEdit;
    FCptPoint: TEdit;
    FCPTLETT: TEdit;
    FMonnaie: TCheckBox;
    Sauve2: TSaveDialog;
    FRazSit: TCheckBox;
    FRaz: TCheckBox;
    Patience: TLabel;
    FRazCli: TCheckBox;
    GP3: TGroupBox;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    FMasquePrefFou: TEdit;
    FRepFou: TEdit;
    FMasqueSuffFou: TEdit;
    BDirFou: TButton;
    FOk3: TCheckBox;
    EBStop: TToolbarButton97;
    FRazFou: TCheckBox;
    Label17: TLabel;
    FRep: THValComboBox;
    procedure BDirTRTClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FOk1Click(Sender: TObject);
    procedure BRecupSISCOClick(Sender: TObject);
    procedure BZoomCptLettClick(Sender: TObject);
    procedure BZoomCptPointClick(Sender: TObject);
    procedure BZoomJalVenClick(Sender: TObject);
    procedure BZoomJalAchClick(Sender: TObject);
    procedure FStatChange(Sender: TObject);
    procedure CBAnaChaClick(Sender: TObject);
    procedure EBStopClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
    { Déclarations privées }
    Recup : Boolean ;
    TOBTRF : TOB ;
    Procedure AlimEcran ;
    procedure SauveEcran(RAZLettrage : Boolean) ;
    Procedure ActiveCol(GB : tGroupBox ; OkOk : Boolean) ;
    Function  OkRecupTRT(Var NomFicTrouve : String) : Integer ;
    Function  OkRecupImmo(Var NomFicTrouve : String) : Integer ;
    Function  OkRecupSit(Var NomFicTrouve : String) : Integer ;
    Function  OkRecupCli(Var NomFicTrouve : String) : Integer ;
    Function  OkRecupFou(Var NomFicTrouve : String) : Integer ;
    Function  WhatFichier : String ;
    Procedure RecupSit(St : String) ;
    Function  TRTCorrect(StFichier : String) : Integer ;
    procedure UpdateOK6;
  public
    { Déclarations publiques }
  end;

implementation

uses ImpAutoS3, LETTREF, MulLettRef, InfFicSI, impFicU, CptSISCO, JalSISCO,ImpCegid, ImEnt,
     ImpExpBDS, RecupII ;

{$R *.DFM}

Procedure RecupSISCOS3(Recup : Boolean) ;
var XX:TFParamRecupSISCO ;
    PP : THPanel ;
BEGIN
XX:=TFParamRecupSISCO.Create(Application) ;
XX.Recup:=Recup ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
   try
     XX.ShowModal ;
     finally
     XX.Free ;
     END ;
   SourisNormale ;
   END else
   BEGIN
   InitInside(XX,PP) ;
   XX.Show ;
   END ;
END ;

Function RetoucheRep(Rep : String) : String ;
BEGIN
Result:=Trim(Rep) ; If Trim(Rep)='' Then Exit ;
If Rep[Length(Rep)]<>'\' Then Rep:=Rep+'\' ;
Result:=Rep ;
END ;

procedure TFParamRecupSISCO.BDirTRTClick(Sender: TObject);
var st : AnsiString ;
    i : integer ;
    Options: TSelectDirOpts ;
    TE,TM,TN  : TEdit ;
    TB  : tButton ;
    LePath,LeNom,Lextension : String ;
begin
TB:=tButton(Sender) ; If TB=Nil Then Exit ;
TM := nil; TN := nil;
Case TB.Tag Of
  0 : TE:=FRepTRT ;
  1 : BEGIN TE:=FRepCli ; TM:=FMasquePrefCli ; TN:=FMasqueSuffCli ; END ;
  2 : BEGIN TE:=FRepFou ; TM:=FMasquePrefFou ; TN:=FMasqueSuffFou ; END ;
  3 : BEGIN TE :=FRepImmo ; TM:=FMasquePrefImmo ; TN:=FMasqueSuffImmo ; END ;
  4 : BEGIN TE :=FRepSit ; TM:=FMasquePrefSit ; TN:=FMasqueSuffSit ; End ;
  Else Exit ;
  END ;
TE.Text:=RetoucheRep(TE.Text) ;
If directoryExists(TE.Text) then St:=TE.text else st:='c:\' ;
i:=1 ; Options:=[] ;
If TB.Tag=0 Then
  BEGIN
  if selectdirectory(st,Options,i) then Te.text:=St ;
  END Else
  BEGIN
  St:=TE.TExt+TM.Text+'.'+TN.Text ;
  DirDefault(Sauve2,St) ;
  if Sauve2.Execute then
    BEGIN
    St:=Sauve2.FileName ;
    LePath:=ExtractFilePath(St) ;
    LExtension:=ExtractFileExt(St) ;
    LeNom:=ExtractFileName(St) ;
    i:=Pos('.',LeNom) ;
    If i>0 Then BEGIN LeNom:=Copy(LeNom,1,i-1) ; Delete(LExtension,1,1) ; END ;
    TE.Text:=LePath ; TM.Text:=LeNom ; TN.Text:=LExtension ;
    END ;
  END ;
end;

Function PasDeFichier(Rep,pref,Suff : String ; Var NomFicTrouve : String) : Boolean ;
Var SearchRec : TSearchRec ;
    LeNom : String ;
BEGIN
NomFicTrouve:='' ;
Result:=TRUE ; If pref='' Then BEGIN Result:=TRUE ; Exit ; END ;
Pref:=Pref+'*' ; If Suff='' Then Suff:='*' ;
LeNom:=Rep+pref+'.'+Suff ;
If FindFirst(LeNom, 0, SearchRec)=0 Then BEGIN Result:=FALSE ; NomFicTrouve:=SearchRec.Name ; END ;
FindClose(SearchRec);
END ;

Function TFParamRecupSISCO.OkRecupTRT(Var NomFicTrouve : String) : Integer ;
BEGIN
NomFicTrouve:='' ;
Result:=0 ;
If Not DirectoryExists(FRepTrt.text) Then
  BEGIN
  Result:=5 ; Pages.activePage:=PStandards ; FRepTRT.setFocus ; Exit ;
  END ;
If PasDeFichier(FRepTrt.text,FmasquePrefTRT.Text,FmasqueSuffTRT.Text,NomFicTrouve) Then
  BEGIN
  Result:=6 ; Pages.activePage:=PStandards ; FmasquePrefTRT.setFocus ; Exit ;
  END ;
If FOk6.Checked Then
  BEGIN
  If (LgGen.Value>0) And (LgGen.Value<5) Then
    BEGIN
    Result:=8 ; Pages.activePage:=PDivers ; LgGen.setFocus ; Exit ;
    END ;
  If (LgAux.Value>0) And (LgAux.Value<5) Then
    BEGIN
    Result:=9 ; Pages.activePage:=PDivers ; LgAux.setFocus ; Exit ;
    END ;
  If (LgAna.Value>0) And (LgAna.Value<5) Then
    BEGIN
    Result:=10 ; Pages.activePage:=PDivers ; LgAna.setFocus ; Exit ;
    END ;
  If BourreGen.Text='' Then BourreGen.Text:='0' ;
  If BourreAux.Text='' Then BourreAux.Text:='0' ;
  If BourreAna.Text='' Then BourreAna.Text:='0' ;
  END ;
If Trim(FJalAn.Text)='' Then
  If (LgAux.Value>0) And (LgAux.Value<5) Then
    BEGIN
    Result:=9 ; Pages.activePage:=PComptes ; FJalAN.setFocus ; Exit ;
    END ;

END ;

Function TFParamRecupSISCO.OkRecupImmo(Var NomFicTrouve : String) : Integer ;
BEGIN
NomFicTrouve:='' ;
Result:=0 ;
If Not DirectoryExists(FRepImmo.text) Then
  BEGIN
  Result:=14 ; Pages.activePage:=PSupp ; FRepImmo.setFocus ; Exit ;
  END ;
If PasDeFichier(FRepImmo.text,FmasquePrefImmo.Text,FmasqueSuffImmo.Text,NomFicTrouve) Then
  BEGIN
  Result:=15 ; Pages.activePage:=PSupp ; FmasquePrefImmo.setFocus ; Exit ;
  END ;
END ;

Function TFParamRecupSISCO.OkRecupSit(Var NomFicTrouve : String) : Integer ;
BEGIN
NomFicTrouve:='' ;
Result:=0 ;
If Not DirectoryExists(FRepSit.text) Then
  BEGIN
  Result:=16 ; Pages.activePage:=PSupp ; FRepSit.setFocus ; Exit ;
  END ;
If PasDeFichier(FRepSit.text,FmasquePrefSit.Text,FmasqueSuffSit.Text,NomFicTrouve) Then
  BEGIN
  Result:=17 ; Pages.activePage:=PSupp ; FmasquePrefImmo.setFocus ; Exit ;
  END ;
END ;

Function TFParamRecupSISCO.OkRecupCli(Var NomFicTrouve : String) : Integer ;
BEGIN
NomFicTrouve:='' ;
Result:=0 ;
If Not DirectoryExists(FRepCli.text) Then
  BEGIN
  Result:=18 ; Pages.activePage:=PStandards ; FRepCli.setFocus ; Exit ;
  END ;
If PasDeFichier(FRepCli.text,FmasquePrefCli.Text,FmasqueSuffCli.Text,NomFicTrouve) Then
  BEGIN
  Result:=19 ; Pages.activePage:=PStandards ; FmasquePrefCli.setFocus ; Exit ;
  END ;
END ;

Function TFParamRecupSISCO.OkRecupFou(Var NomFicTrouve : String) : Integer ;
BEGIN
NomFicTrouve:='' ;
Result:=0 ;
If Not DirectoryExists(FRepFou.text) Then
  BEGIN
  Result:=20 ; Pages.activePage:=PSupp ; FRepFou.setFocus ; Exit ;
  END ;
If PasDeFichier(FRepFou.text,FmasquePrefFou.Text,FmasqueSuffFou.Text,NomFicTrouve) Then
  BEGIN
  Result:=21 ; Pages.activePage:=PSupp ; FmasquePrefFou.setFocus ; Exit ;
  END ;
END ;

Procedure TFParamRecupSISCO.RecupSit(St : String) ;
var Import : TImpExpBDS;
BEGIN
Import := nil;
Patience.Visible:=TRUE ; Application.ProcessMEssages ;
If FRAZSit.Checked Then
  BEGIN
  ExecuteSQL('DELETE FROM CBALSIT WHERE BSI_CODEBAL LIKE "SISCO%" ') ;
  ExecuteSQL('DELETE FROM CBALSITECR WHERE BSE_CODEBAL LIKE "SISCO%" ') ;
  END ;
try
  Import := TImpExpBDS.Create ( miSiscoII ) ;
  Import.OnInformation := Nil;
  Import.Importation ( St , 'SISCO', ' Balance SISCO récupérée' ,'SISCO' ,TRUE);
  Patience.Visible:=FALSE ; Application.ProcessMEssages ;
  if Import.GetLastError = 0 then  // Tout est OK
  begin
    PGIInfo ('Importation terminée avec succès.','Récupération Balance Sisco II');
  end else PGIBox ( 'Erreur lors de la récupération Balance Sisco II.','');
finally
  Import.Free;
End ;
END ;

Function TFParamRecupSISCO.TRTCorrect(StFichier : String) : Integer ;
Var Fichier : TextFile ;
    St : String ;
    Ok00,Ok03,Ok50 : Boolean ;
    Monnaie,SLgCpt,SLgAna : String ;
    LgCpt,LgSect : Integer ;

BEGIN
Result:=0 ;
SLgCpt := '0'; SLgAna := '0';
If Not FOK6.Checked Then Exit ;
AssignFile(Fichier,StFichier) ; {$i-} Reset(Fichier) ; {$i+} If IOresult<>0 THen Exit ;
If Not Eof(Fichier) THen
  BEGIN
  Readln(Fichier,St) ; If Copy(St,1,11)<>'***DEBUT***' Then BEGIN Result:=23 ; Exit ; END ;
  END ;
Ok00:=FALSE ; Ok03:=FALSE ; Ok50:=FALSE ;
LgCpt:=4 ; LgSect:=4 ;
While Not Eof(Fichier) Do
  BEGIN
  ReadLn(Fichier,St) ;
  If (Copy(St,1,2)='00') Then BEGIN Ok00:=TRUE ; SLgCpt:=Trim(Copy(St,21,2)) ; END ;
  If (Copy(St,1,2)='03') Then
    BEGIN
    Ok03:=TRUE ;
    If Copy(St,54,1)='N' Then BEGIN Ok50:=TRUE ; LgSect:=-1 ; END ;
    END ;
  If (Copy(St,1,2)='50') Then
    BEGIN
    Ok50:=TRUE ;
    SLgAna:=Copy(St,44,2) ;
    END ;
  If Ok00 And Ok03 And Ok50 Then Break ;
  END ;
System.Close(Fichier) ;
LgCpt:=StrToInt(SLgCpt) ;
If LgGen.Value<LgCpt Then BEGIN Result:=24 ; Exit ; END ;
If LgAux.Value<LgCpt Then BEGIN Result:=25 ; Exit ; END ;
If LgSect<>-1 Then
  BEGIN
  LgSect:=StrToInt(SLgAna) ;
  If LgAna.Value<LgSect Then BEGIN Result:=27 ; Exit ; END ;
  END ;
END ;

procedure TFParamRecupSISCO.BValiderClick(Sender: TObject);
Var Rep,Fic,Mon,Ana,TRF,Decoupe,RAZ,Stat : String ;
    i : Integer ;
    SiscoEstRecupere : Boolean ;
    OkSauve,PbRecup : Boolean ;
    OkTRT,OkCLI,OkFou,OkImmo,OkSit : Boolean ;
    IndRep : Integer ;
    NomTRTTrouve,RienABranler : String ;
begin
If Recup Then If Msg.execute(0,Caption,'')<>mrYes Then exit ;
Rep:=FRepTRT.Text ; FRepTRT.Text:=RetoucheRep(Rep) ;
Rep:=FRepImmo.Text ; FRepImmo.Text:=RetoucheRep(Rep) ;
Rep:=FRepSit.Text ; FRepSit.Text:=RetoucheRep(Rep) ;
Rep:=FRepCli.Text ; FRepCli.Text:=RetoucheRep(Rep) ;
Rep:=FRepFou.Text ; FRepFou.Text:=RetoucheRep(Rep) ;
SiscoEstRecupere:=FALSE ;
If Recup Then
  BEGIN
  SauveEcran(TRUE) ;
  PbRecup:=FALSE ;
  OkTRT:=TRUE ; OkCLI:=TRUE ; OkFou:=TRUE ; OkImmo:=TRUE ; OkSit:=TRUE ;
  If FOk1.Checked Then BEGIN i:=OkRecupTRT(NomTrtTrouve) ; OkTRT:=i=0 ; If Not OkTRT Then BEGIN Msg.execute(i,Caption,'') ; Exit ; END ; END ;
  If FOk2.Checked Then BEGIN i:=OkRecupCli(RienABranler) ; OkCli:=i=0 ; If Not OkCli Then BEGIN Msg.execute(i,Caption,'') ; Exit ; END ; END ;
  If FOk3.Checked Then BEGIN i:=OkRecupFou(RienABranler) ; OkFou:=i=0 ; If Not OkFou Then BEGIN Msg.execute(i,Caption,'') ; Exit ; END ; END ;
  If FOk4.Checked Then BEGIN i:=OkRecupImmo(RienABranler) ; OkImmo:=i=0 ; If Not OkImmo Then BEGIN Msg.execute(i,Caption,'') ; Exit ; END ; END ;
  If FOk5.Checked Then BEGIN i:=OkRecupSit(RienABranler) ; OkSit:=i=0 ; If Not OkSit Then BEGIN Msg.execute(i,Caption,'') ; Exit ; END ; END ;
  If FOk1.Checked Then
    BEGIN
    VH^.STOPRSP:=FALSE ;
    If OkTrt Then
      BEGIN
      i:=TRTCorrect(FRepTRT.Text+'\'+NomTRTTrouve) ;
      If i>0 Then BEGIN Msg.execute(i,'','') ; Exit ; END ;
      If Not FMonnaie.Checked Then
        BEGIN
        If Msg.execute(26,'','')<>mrYes Then Exit ;
        END ;
      If i<=0 Then
        BEGIN
        Rep:=FRepTrt.text ; Fic:=FMasquePrefTRT.Text+'.'+FMasqueSuffTRT.text ;
        If FMonnaie.Checked Then Mon:='E' Else Mon:='F' ;
        Ana:='O' ; TRF:='000' ;
        If FBigVol.Checked Then Decoupe:='X' Else Decoupe:='-' ;
        If FRAZ.Checked Then RAZ:='X' Else RAZ:='-' ;
        Stat:=FStat.Value ;
        SiscoEstRecupere:=ImportPgeS3(Rep,Fic,Mon,Ana,TRF,Decoupe,RAZ,Stat) ;
        END ;
      END ;
    END ;
  If FOk2.Checked Then
    BEGIN
    VH^.STOPRSP:=FALSE ;
    If OkCli Then
      BEGIN
      Rep:=FRepCli.text ; Fic:=FMasquePrefCli.Text+'.'+FMasqueSuffCli.text ;
      Fic:=Rep+Fic ;
      Patience.Visible:=TRUE ; Application.ProcessMEssages ;
      If FRazCli.Checked Then
        BEGIN
        ExecuteSQL('DELETE FROM RIB WHERE R_SOCIETE="1"') ; ExecuteSQL('DELETE FROM CONTACT WHERE C_TEXTELIBRE3="1"') ;
        ExecuteSQL('DELETE FROM NATCPTE WHERE (NT_TYPECPTE="T00" OR NT_TYPECPTE="T01" OR NT_TYPECPTE="T02") And NT_TEXTE9="1" ') ;
        ExecuteSQL('DELETE FROM MODEREGL WHERE MR_SOCIETE="1"') ;
        END ;
      IndRep:=-1 ; If FRep.ItemIndex>0 Then IndRep:=FRep.ItemIndex-1 ;
      EBStop.Visible:=TRUE ; Application.ProcessMessages ;
      RecupSuiviClient(Fic,IndRep) ;
      EBStop.Visible:=FALSE ; Patience.Visible:=FALSE ; Application.ProcessMEssages ;
      PGIInfo ('Importation terminée avec succès.','Récupération Suivi client II');
      END ;
    END ;
  If FOk3.Checked Then
    BEGIN
    VH^.STOPRSP:=FALSE ;
    If OkFou Then
      BEGIN
      Rep:=FRepFou.text ; Fic:=FMasquePrefFou.Text+'.'+FMasqueSuffFou.text ;
      Fic:=Rep+Fic ;
      Patience.Visible:=TRUE ; Application.ProcessMEssages ;
      If FRazFou.Checked Then
        BEGIN
        ExecuteSQL('DELETE FROM RIB WHERE R_SOCIETE="2"') ; ExecuteSQL('DELETE FROM CONTACT WHERE C_TEXTELIBRE3="2"') ;
        ExecuteSQL('DELETE FROM NATCPTE WHERE (NT_TYPECPTE="T00" OR NT_TYPECPTE="T01" OR NT_TYPECPTE="T02") And NT_TEXTE9="2" ') ;
        ExecuteSQL('DELETE FROM MODEREGL WHERE MR_SOCIETE="2"') ;
        END ;
      EBStop.Visible:=TRUE ; Application.ProcessMessages ;
      RecupSuiviFournisseur(Fic) ;
      EBStop.Visible:=FALSE ; Patience.Visible:=FALSE ; Application.ProcessMEssages ;
      PGIInfo ('Importation terminée avec succès.','Récupération Suivi fournisseur II');
      END ;
    END ;
  If FOk4.Checked Then
    BEGIN
    VH^.STOPRSP:=FALSE ;
    If OkImmo Then
      BEGIN
      Rep:=FRepImmo.text ; Fic:=FMasquePrefImmo.Text+'.'+FMasqueSuffImmo.text ;
      Fic:=Rep+Fic ;
      ChargeVHImmo ;
      SavImo2PGI(Fic,TRUE);
      PGIInfo ('Importation terminée avec succès.','Récupération Imo II');
      END ;
    END ;
  If FOk5.Checked Then
    BEGIN
    VH^.STOPRSP:=FALSE ;
    If OkSit Then
      BEGIN
      Rep:=FRepSit.text ; Fic:=FMasquePrefSit.Text+'.'+FMasqueSuffSit.text ;
      Fic:=Rep+Fic ;
      RecupSit(Fic) ;
      END ;
    END ;
  END ;
If Not Recup Then
  BEGIN
  OkSauve:=FALSE ;
  If FOk1.Checked Then BEGIN OkSauve:=TRUE ; i:=OkRecupTRT(RienABranler) ; If i<>0 Then Msg.execute(i,Caption,'') ; END ;
  If FOk2.Checked Then BEGIN OkSauve:=TRUE ; i:=OkRecupCli(RienABranler) ; If i<>0 Then Msg.execute(i,Caption,'') ; END ;
  If FOk3.Checked Then BEGIN OkSauve:=TRUE ; i:=OkRecupFou(RienABranler) ; If i<>0 Then Msg.execute(i,Caption,'') ; END ;
  If FOk4.Checked Then BEGIN OkSauve:=TRUE ; i:=OkRecupImmo(RienABranler) ; If i<>0 Then Msg.execute(i,Caption,'') ; END ;
  If FOk5.Checked Then BEGIN OkSauve:=TRUE ; i:=OkRecupSit(RienABranler) ; If i<>0 Then Msg.execute(i,Caption,'') ; END ;
  If OkSauve And (Msg.Execute(4,Caption,'')=mrYes) Then SauveEcran(FALSE) ;
  END ;
end;

procedure TFParamRecupSISCO.AlimEcran ;
Var TobL : TOB ;
    i : Integer ;
    NomChamp : String ;
    ValChamp : String ;
BEGIN
If TobTRF.Detail.Count<=0 Then Exit ;
For i:=0 To TobTRF.Detail.Count-1 Do
  BEGIN
  TobL:=TobTRF.Detail[i] ;
  NomChamp:=TOBL.GetValue('TRP_NOM') ;
  ValChamp:=TOBL.GetValue('TRP_DATA') ;
  If NomChamp='CP_NOMPREFTRT' Then FMasquePrefTRT.text:=ValChamp Else
  If NomChamp='CP_NOMSUFFTRT' Then FMasqueSuffTRT.text:=ValChamp Else
  If NomChamp='CP_REPERTOIRETRT' Then FRepTRT.text:=ValChamp Else
  If NomChamp='CP_NOMPREFCLI' Then FMasquePrefCli.text:=ValChamp Else
  If NomChamp='CP_NOMSUFFCLI' Then FMasqueSuffCli.text:=ValChamp Else
  If NomChamp='CP_REPERTOIRECLI' Then FRepCli.text:=ValChamp Else
  If NomChamp='CP_NOMPREFFOU' Then FMasquePrefFou.text:=ValChamp Else
  If NomChamp='CP_NOMSUFFOU' Then FMasqueSuffFou.text:=ValChamp Else
  If NomChamp='CP_REPERTOIREFOU' Then FRepFou.text:=ValChamp Else
  If NomChamp='CP_NOMPREFIMMO' Then FMasquePrefImmo.text:=ValChamp Else
  If NomChamp='CP_NOMSUFFIMMO' Then FMasqueSuffImmo.text:=ValChamp Else
  If NomChamp='CP_REPERTOIREIMMO' Then FRepImmo.text:=ValChamp Else
  If NomChamp='CP_NOMPREFSIT' Then FMasquePrefSit.text:=ValChamp Else
  If NomChamp='CP_NOMSUFFSIT' Then FMasqueSuffSit.text:=ValChamp Else
  If NomChamp='CP_REPERTOIRESIT' Then FRepSit.text:=ValChamp Else
  If NomChamp='CP_CPTELETTRABLE' Then FCptLett.text:=ValChamp Else
  If NomChamp='CP_CPTEPOINTABLE' Then FCptPoint.text:=ValChamp Else
  If NomChamp='CP_JOURNALAN' Then FJalAN.text:=ValChamp Else
  If NomChamp='CP_JALVENTE' Then FjalVen.text:=ValChamp Else
  If NomChamp='CP_JALACHAT' Then FJalAch.text:=ValChamp Else
  If NomChamp='CP_CPTEECARTCONV' Then FCptEcart.text:=ValChamp Else
  If NomChamp='CP_LGGEN' Then BEGIN If Trim(ValChamp)='' Then LgGen.Value:=0 Else LgGen.Value:=StrToInt(ValChamp) ; END Else
  If NomChamp='CP_LGAUX' Then BEGIN If Trim(ValChamp)='' Then LgAux.Value:=0 Else LgAux.Value:=StrToInt(ValChamp) ; END Else
  If NomChamp='CP_LGANA' Then BEGIN If Trim(ValChamp)='' Then LgAna.Value:=0 Else LgAna.Value:=StrToInt(ValChamp) ; END Else
  If NomChamp='CP_BOURREGEN' Then BEGIN If ValChamp='' Then BourreGen.text:='0' Else BourreGen.text:=ValChamp ; END Else
  If NomChamp='CP_BOURREAUX' Then BEGIN If ValChamp='' Then BourreAux.text:='0' Else BourreAux.text:=ValChamp END Else
  If NomChamp='CP_BOURREANA' Then BEGIN If ValChamp='' Then BourreAna.text:='0' Else BourreAna.text:=ValChamp END Else
  If NomChamp='CP_CPTAEURO' Then FMonnaie.Checked:=ValChamp='X' Else
  If NomChamp='CP_TRTOK' Then FOk1.Checked:=ValChamp='X' Else
  If NomChamp='CP_CLIOK' Then FOk2.Checked:=ValChamp='X' Else
  If NomChamp='CP_FOUOK' Then FOk3.Checked:=ValChamp='X' Else
  If NomChamp='CP_IMMOOK' Then FOk4.Checked:=ValChamp='X' Else
  If NomChamp='CP_SITOK' Then FOk5.Checked:=ValChamp='X' Else
  If NomChamp='CP_RAZ' Then FRAZ.Checked:=ValChamp='X' Else
  If NomChamp='CP_RAZSIT' Then FRAZSit.Checked:=ValChamp='X' Else
  If NomChamp='CP_RAZCLI' Then FRAZCli.Checked:=ValChamp='X' Else
  If NomChamp='CP_RAZFOU' Then FRAZFou.Checked:=ValChamp='X' Else
  If NomChamp='CP_RECUPANACHA' Then CbAnaCha.Checked:=ValChamp='X' Else
  If NomChamp='CP_RECUPANAPRO' Then CbAnaPro.Checked:=ValChamp='X' Else
  If NomChamp='CP_LGOK' Then FOk6.Checked:=ValChamp='X' Else
  If NomChamp='CP_BIGVOL' Then FBigVol.Checked:=ValChamp='X' Else
  If NomChamp='CP_STAT' Then FStat.Value:=ValChamp Else
  If NomChamp='CP_INDREP' Then BEGIN If Trim(ValChamp)='' Then FRep.ItemIndex:=0 Else FRep.ItemIndex:=StrToInt(ValChamp) ; END Else
  If NomChamp='CP_LETTRAGE' Then FTYPELettrage.Value:=ValChamp ;
  END ;
If FTYPELettrage.ItemIndex<0 Then FTYPELettrage.ItemIndex:=0 ;
If FStat.ItemIndex<0 Then FStat.ItemIndex:=0 ;
If fJalAN.text='' Then fJalAN.Text:='aa;' ;
END ;

procedure TFParamRecupSISCO.SauveEcran(RAZLettrage : Boolean) ;
Var TobL : TOB ;
    i : Integer ;
    NomChamp : String ;
BEGIN
If TobTRF.Detail.Count<=0 Then Exit ;
For i:=0 To TobTRF.Detail.Count-1 Do
  BEGIN
  TobL:=TobTRF.Detail[i] ;
  NomChamp:=TOBL.GetValue('TRP_NOM') ;
  If NomChamp='CP_NOMPREFTRT' Then TOBL.PutValue('TRP_DATA',FMasquePrefTRT.text) Else
  If NomChamp='CP_NOMSUFFTRT' Then TOBL.PutValue('TRP_DATA',FMasqueSuffTRT.text) Else
  If NomChamp='CP_REPERTOIRETRT' Then TOBL.PutValue('TRP_DATA',FRepTRT.text) Else
  If NomChamp='CP_NOMPREFCLI' Then TOBL.PutValue('TRP_DATA',FMasquePrefCli.text) Else
  If NomChamp='CP_NOMSUFFCLI' Then TOBL.PutValue('TRP_DATA',FMasqueSuffCli.text) Else
  If NomChamp='CP_REPERTOIRECLI' Then TOBL.PutValue('TRP_DATA',FRepCli.text) Else
  If NomChamp='CP_NOMPREFFOU' Then TOBL.PutValue('TRP_DATA',FMasquePrefFou.text) Else
  If NomChamp='CP_NOMSUFFOU' Then TOBL.PutValue('TRP_DATA',FMasqueSuffFou.text) Else
  If NomChamp='CP_REPERTOIREFOU' Then TOBL.PutValue('TRP_DATA',FRepFou.text) Else
  If NomChamp='CP_NOMPREFIMMO' Then TOBL.PutValue('TRP_DATA',FMasquePrefImmo.text) Else
  If NomChamp='CP_NOMSUFFIMMO' Then TOBL.PutValue('TRP_DATA',FMasqueSuffImmo.text) Else
  If NomChamp='CP_REPERTOIREIMMO' Then TOBL.PutValue('TRP_DATA',FRepImmo.text) Else
  If NomChamp='CP_NOMPREFSIT' Then TOBL.PutValue('TRP_DATA',FMasqueprefSit.text) Else
  If NomChamp='CP_NOMSUFFSIT' Then TOBL.PutValue('TRP_DATA',FMasqueSuffSit.text) Else
  If NomChamp='CP_REPERTOIRESIT' Then TOBL.PutValue('TRP_DATA',FRepSit.text) Else
  If NomChamp='CP_CPTELETTRABLE' Then TOBL.PutValue('TRP_DATA',FCptLett.text) Else
  If NomChamp='CP_CPTEPOINTABLE' Then TOBL.PutValue('TRP_DATA',FCptPoint.text) Else
  If NomChamp='CP_JOURNALAN' Then TOBL.PutValue('TRP_DATA',FJalAN.text) Else
  If NomChamp='CP_JALVENTE' Then TOBL.PutValue('TRP_DATA',FjalVen.text) Else
  If NomChamp='CP_JALACHAT' Then TOBL.PutValue('TRP_DATA',FJalAch.text) Else
  If NomChamp='CP_CPTEECARTCONV' Then TOBL.PutValue('TRP_DATA',FCptEcart.text) Else
  If NomChamp='CP_LGGEN' Then TOBL.PutValue('TRP_DATA',IntToStr(LgGen.Value)) Else
  If NomChamp='CP_LGAUX' Then TOBL.PutValue('TRP_DATA',IntToStr(LgAux.Value)) Else
  If NomChamp='CP_LGANA' Then TOBL.PutValue('TRP_DATA',IntToStr(LgAna.Value)) Else
  If NomChamp='CP_BOURREGEN' Then TOBL.PutValue('TRP_DATA',BourreGen.Text) Else
  If NomChamp='CP_BOURREAUX' Then TOBL.PutValue('TRP_DATA',BourreAux.Text) Else
  If NomChamp='CP_BOURREANA' Then TOBL.PutValue('TRP_DATA',BourreAna.Text) Else
  If NomChamp='CP_CPTAEURO' Then BEGIN If FMonnaie.Checked Then TOBL.PutValue('TRP_DATA','X') Else TOBL.PutValue('TRP_DATA','-') ; End Else
  If NomChamp='CP_TRTOK' Then BEGIN If FOk1.Checked Then TOBL.PutValue('TRP_DATA','X') Else TOBL.PutValue('TRP_DATA','-') ; End Else
  If NomChamp='CP_CLIOK' Then BEGIN If FOk2.Checked Then TOBL.PutValue('TRP_DATA','X') Else TOBL.PutValue('TRP_DATA','-') ; End Else
  If NomChamp='CP_FOUOK' Then BEGIN If FOk3.Checked Then TOBL.PutValue('TRP_DATA','X') Else TOBL.PutValue('TRP_DATA','-') ; End Else
  If NomChamp='CP_IMMOOK' Then BEGIN If FOk4.Checked Then TOBL.PutValue('TRP_DATA','X') Else TOBL.PutValue('TRP_DATA','-') ; End Else
  If NomChamp='CP_SITOK' Then BEGIN If FOk5.Checked Then TOBL.PutValue('TRP_DATA','X') Else TOBL.PutValue('TRP_DATA','-') ; End Else
  If NomChamp='CP_RAZ' Then BEGIN If FRAZ.Checked Then TOBL.PutValue('TRP_DATA','X') Else TOBL.PutValue('TRP_DATA','-') ; End Else
  If NomChamp='CP_RAZSIT' Then BEGIN If FRAZSit.Checked Then TOBL.PutValue('TRP_DATA','X') Else TOBL.PutValue('TRP_DATA','-') ; End Else
  If NomChamp='CP_RAZCLI' Then BEGIN If FRAZCli.Checked Then TOBL.PutValue('TRP_DATA','X') Else TOBL.PutValue('TRP_DATA','-') ; End Else
  If NomChamp='CP_RAZFOU' Then BEGIN If FRAZFou.Checked Then TOBL.PutValue('TRP_DATA','X') Else TOBL.PutValue('TRP_DATA','-') ; End Else
  If NomChamp='CP_RECUPANACHA' Then BEGIN If CBAnaCHA.Checked Then TOBL.PutValue('TRP_DATA','X') Else TOBL.PutValue('TRP_DATA','-') ; End Else
  If NomChamp='CP_RECUPANAPRO' Then BEGIN If CBAnaPRO.Checked Then TOBL.PutValue('TRP_DATA','X') Else TOBL.PutValue('TRP_DATA','-') ; End Else
  If NomChamp='CP_LGOK' Then BEGIN If FOk6.Checked Then TOBL.PutValue('TRP_DATA','X') Else TOBL.PutValue('TRP_DATA','-') ; End Else
  If NomChamp='CP_BIGVOL' Then BEGIN If FBigVol.Checked Then TOBL.PutValue('TRP_DATA','X') Else TOBL.PutValue('TRP_DATA','-') ; End Else
  If NomChamp='CP_STAT' Then TOBL.PutValue('TRP_DATA',FStat.Value) ;
  If NomChamp='CP_INDREP' Then TOBL.PutValue('TRP_DATA',IntToStr(FRep.ItemIndex)) Else
  If NomChamp='CP_LETTRAGE' Then TOBL.PutValue('TRP_DATA',FTYPELettrage.Value) ;
  If RAZLettrage Then
    BEGIN
    If NomChamp='CP_NBTIERSL' Then TOBL.PutValue('TRP_DATA','0') Else
    If NomChamp='CP_NBGENEL' Then TOBL.PutValue('TRP_DATA','0') Else
    If NomChamp='CP_TIERSL' Then TOBL.PutValue('TRP_DATA','') Else
    If NomChamp='CP_GENEL' Then TOBL.PutValue('TRP_DATA','') ;
    END ;
  END ;
TobTRF.UpdateDB ;
END ;


procedure TFParamRecupSISCO.FormShow(Sender: TObject);
Var Q : TQuery ;
begin
Pages.activePage:=PStandards ; FOk1.setFocus ;
Q:=OpenSQL('SELECT * FROM TRFPARAM WHERE TRP_CODE="000" ',TRUE) ;
TOBTRF.LoadDetailDB('TRFPARAM','','',Q,False,True) ;
Ferme(Q) ;
AlimEcran ;
Fok1Click(FOk1) ; Fok1Click(FOk2) ; Fok1Click(FOk3) ; Fok1Click(FOk4) ; Fok1Click(FOk5) ; Fok1Click(FOk6) ;
end;

procedure TFParamRecupSISCO.FormCreate(Sender: TObject);
begin
TOBTRF:=TOB.Create('',Nil,-1) ;
end;

procedure TFParamRecupSISCO.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
TobTRF.Free ;
end;

Procedure TFParamRecupSISCO.ActiveCol(GB : tGroupBox ; OkOk : Boolean) ;
var I : Integer;
    ChildControl: TControl;
begin
If GB=Nil Then Exit ;
for I:= 0 to GB.ControlCount -1 do
  BEGIN
  ChildControl:=GB.Controls[I];
  If ChildControl<>NIl Then
    BEGIN
    ChildControl.Enabled:=OkOk ;
    END ;
  END ;
end;

procedure TFParamRecupSISCO.FOk1Click(Sender: TObject);
Var St : String ;
    CB : tCheckBox ;
    GB : TGroupBox ;
begin
CB:=TCheckBox(Sender) ; If CB=NIL Then Exit ;
If CB.Tag>0 Then
  BEGIN
  GB:=TGroupBox(FindComponent('GP'+IntToStr(CB.Tag))) ; If GB<>NIL Then ActiveCol(GB,CB.Checked) ;
  END ;
Cb.Enabled:=TRUE ;
if (CB = FOK6) then UpdateOK6;
end;

Function IsFichierUnique(StFichier : String) : Boolean ;
BEGIN
Result:=FALSE ;
If (Pos('*',StFichier)>0) Or (Pos('?',StFichier)>0) Then Exit ;
Result:=FileExists(StFichier) ;
END ;

Function TFParamRecupSISCO.WhatFichier : String ;
Var Fichier : TextFile ;
    St : String ;
    Ok00,Ok01,Ok02,Ok03,Ok50 : Boolean ;
    Rep,pref,Suff : String ;
    StFichier,StFichierUnique : String ;
    OkOk : Boolean ;
BEGIN
Result:='' ;
Rep:=FRepTRT.Text ; Rep:=RetoucheRep(Rep) ;
Pref:=FMasqueprefTRT.Text ;
Suff:=FMasqueSuffTRT.Text ;
StFichierUnique:=Rep+Pref+'.'+Suff ;
If IsFichierUnique(StFichierUnique) Then OkOk:=TRUE Else
  BEGIN
  StFichierUnique:='' ;
  If pref<>'' Then St:=Rep+Pref+'*.'+Suff Else St:=Rep ;
  DirDefault(Sauve,St) ;
  Okok:=Sauve.Execute ;
  If OkOk Then StFichierUnique:=Sauve.FileName ;
  END ;
if OkOk then
  BEGIN
  Result:=StFichierUnique ;
  If Not EstFichierSISCO(Result,TRUE) Then Result:='' ;
  END ;
END ;

procedure TFParamRecupSISCO.BRecupSISCOClick(Sender: TObject);
Var Fichier : TextFile ;
    St : String ;
    Ok00,Ok01,Ok02,Ok03,Ok50 : Boolean ;
    Rep,pref,Suff : String ;
    StFichier : String ;
begin
StFichier:=WhatFichier ; If StFichier='' Then Exit ;
If InfoFicSISCORecup(StFichier) Then
  BEGIN
  DecomposeNomFic(StFichier,Rep,Pref,Suff) ;
  If Suff<>'' Then
    BEGIN
    pref:=VireTouteSub(pref,Suff) ; Suff:=VireTouteSub(Suff,'.') ;
    END ;
  FRepTRT.Text:=Rep ;
  FMasqueprefTRT.Text:=Pref+'*' ;
  FMasqueSuffTRT.Text:=Suff ;
  END ;
(*
AssignFile(Fichier,StFichier) ; {$i-} Reset(Fichier) ; {$i+}
InitMove(1000,'') ;
Ok00:=FALSE ; Ok01:=FALSE ; Ok02:=FALSE ; Ok03:=FALSE ; Ok50:=FALSE ;
While Not Eof(Fichier) Do
  BEGIN
  MoveCur(FALSE) ;
  ReadLn(Fichier,St) ;
  If (Copy(St,1,2)='00') Then
    BEGIN
    Ok00:=TRUE ;
    _LgGen:=Copy(St,21,2) ;
    _LgAux:=Copy(St,21,2) ;
    END ;
  If (Copy(St,1,2)='01') Then
    BEGIN
    Ok01:=TRUE ;
    END ;
  If (Copy(St,1,2)='02') Then
    BEGIN
    Ok02:=TRUE ;
    END ;
  If (Copy(St,1,2)='03') Then
    BEGIN
    Ok03:=TRUE ;
    If Copy(St,54,1)='N' Then Ok50:=TRUE ;
    END ;
  If (Copy(St,1,2)='50') Then
    BEGIN
    Ok50:=TRUE ;
    _LgAna:=Copy(St,44,2) ;
    END ;
  If Ok00 And Ok01 And Ok02 And Ok03 And Ok50 Then Break ;
  END ;
System.Close(Fichier) ;
FiniMove ;
*)
end;

procedure TFParamRecupSISCO.BZoomCptLettClick(Sender: TObject);
Var Fichier : TextFile ;
    St : String ;
    StFichier : String ;
begin
StFichier:=WhatFichier ; If StFichier='' Then Exit ;
St:=VisuCptSisco(StFichier) ;
If St<>'' Then FCptLett.Text:=St ;
END ;

procedure TFParamRecupSISCO.BZoomCptPointClick(Sender: TObject);
Var Fichier : TextFile ;
    St : String ;
    StFichier : String ;
begin
StFichier:=WhatFichier ; If StFichier='' Then Exit ;
St:=VisuCptSisco(StFichier,1) ;
If St<>'' Then FCptPoint.Text:=St ;
END ;

procedure TFParamRecupSISCO.BZoomJalVenClick(Sender: TObject);
Var Fichier : TextFile ;
    St : String ;
    StFichier : String ;
begin
StFichier:=WhatFichier ; If StFichier='' Then Exit ;
St:=VisuJalSisco2(StFichier) ;
If St<>'' Then FJalVen.Text:=St ;
END ;

procedure TFParamRecupSISCO.BZoomJalAchClick(Sender: TObject);
Var Fichier : TextFile ;
    St : String ;
    StFichier : String ;
begin
StFichier:=WhatFichier ; If StFichier='' Then Exit ;
St:=VisuJalSisco2(StFichier) ;
If St<>'' Then FJalAch.Text:=St ;
END ;

procedure TFParamRecupSISCO.FStatChange(Sender: TObject);
begin
If FStat.Value='ANA' Then
  BEGIN
  If CBAnaCha.Checked Or CBAnaPro.Checked Then BEGIN FStat.Value:='RIE' ; Msg.Execute(12,Caption,'') ; END ;
  END ;
end;

procedure TFParamRecupSISCO.CBAnaChaClick(Sender: TObject);
begin
If (CBAnaCha.Checked Or CBAnaPro.Checked) And (FStat.Value='ANA') Then
  BEGIN
  CBAnaCha.Checked:=FALSE ; CBAnaPro.Checked:=FALSE ; Msg.Execute(13,Caption,'') ;
  END ;
end;

procedure TFParamRecupSISCO.EBStopClick(Sender: TObject);
begin
If Not VH^.STOPRSP Then
  BEGIN
  If Msg.execute(22,Caption,'')=mrYes Then BEGIN VH^.STOPRSP:=TRUE ; Application.ProcessMessages ; End ;
  END ;
end;

procedure TFParamRecupSISCO.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

// FQ 11754
procedure TFParamRecupSISCO.UpdateOK6;
var
  sz : String;
begin
  sz := VH^.Cpta[fbGene].Cb;
  BourreGen.Text := sz; BourreGen.Enabled := not(sz <> '');
  sz := VH^.Cpta[fbAux].Cb;
  BourreAux.Text := sz; BourreAux.Enabled := not(sz <> '');
  sz := VH^.Cpta[fbAxe1].Cb;  // Voir table Axe
  BourreAna.Text := sz; BourreAna.Enabled := not(sz <> '');
end;

end.
