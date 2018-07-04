unit MenuSYNTH;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, DB, HEnt1, Ent1, Hctrls, DBTables, Mask, IniFiles,
  Menus, ComCtrls, HStatus,hmsgbox, MAJSTRUC,MAJTABLE, HPanel,HQry, HTB97,
  PGIExec, PGIEnv;

type
  TMenuManuel = class(TForm)
    Status: THStatusBar;
    mess: THMsgBox;
    Panel1: THPanel;
    BDeconnect: TToolbarButton97;
    BConnect: TToolbarButton97;
    BQuitter: TToolbarButton97;
    HPanel2: THPanel;
    TFDossier: THLabel;
    TFUser: THLabel;
    TFPassword: THLabel;
    FDossier: THValComboBox;
    FUser: TEdit;
    FPassword: TEdit;
    FDateEntree: TMaskEdit;
    TFDateEntree: THLabel;
    Image1: TImage;
    GB1: THPanel;
    BEtatsLibres: TToolbarButton97;
    ImageS5: TImage;
    Timer: TTimer;
    ToolbarButton971: TToolbarButton97;
    ToolbarButton972: TToolbarButton97;
    procedure BAnnulerClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BquitterClick(Sender: TObject);
    procedure BConnectClick(Sender: TObject);
    procedure BDeconnectClick(Sender: TObject);
    procedure FDossierEnter(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure BEtatsLibresClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FDossierClick(Sender: TObject);
    procedure TimerTimer(Sender: TObject);
    procedure ToolbarButton971Click(Sender: TObject);
    procedure ToolbarButton972Click(Sender: TObject);
  private
    SuperUser : Boolean ;
    OkEnvParams : boolean;
    OkVerifStructure : boolean;
    FirstEnvParams : boolean;
    procedure ActiveFonction(Ok : Boolean) ;
    procedure Fermeture;
    procedure GetEnvParams;
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

var MenuManuel : TMenuManuel;

implementation

uses Licutil, CRITSYNTH,CritEdt ;

{$R *.DFM}

procedure TMenuManuel.BAnnulerClick(Sender: TObject);
begin
Close ;
end;

procedure TMenuManuel.FDossierEnter(Sender: TObject);
Var Old : Integer ;
begin
Old:=FDossier.ItemIndex ;
ChargeDossier(FDossier.Items,True,False) ;
if Old<FDossier.Items.Count then FDossier.ItemIndex:=Old ;
end;

procedure TMenuManuel.FormShow(Sender: TObject);
var i : integer ;
begin
  GridPCL:=TRUE ;    
  GB1.Visible := True;
  GB1.Enabled := False;
  VStatus:=Status ;
  ChangeForme(Self) ;
  // ------------------- MD ----------------------
  //# lecture environnement
  If (Not V_PGI.CegidApalatys) then
  begin
    // Création de l'instance unique
    V_PGI_Env := TPGIEnv.Create ;
    //# lit param de PGIApp et màj V_PGI.RunFromLanceur
    InitPGIEnv();
    //# retour au mode APA
    if V_PGI_Env.ModeFonc='APA' then
    begin
      V_PGI_Env.Free;
      V_PGI_Env := Nil;
    end
    //# routage halsocini
    else  If Copy(V_PGI_Env.ModeFonc,1,1)='M' then HalSocIni := 'CEGIDPGI.INI';
  end;
  If V_PGI.RunFromLanceur then V_PGI.MultiUserLogin := True;
  // ------------------- FIN MD --------------------
  If V_PGI.RunFromLanceur then ActiveFonction(True);
  Fermeture; //# appelle ChargeDossier qui lit le ini
  GetEnvParams ;
  //  if OkEnvParams then gConnect.Visible:=FALSE ;
end;

procedure TMenuManuel.GetEnvParams;
var j,i : integer ;
    St,Nom,Value : String ;
    OkUser,OkPass,OkDate,OkDoss : Boolean ;
begin
// exemple /MAJSTRUCTURE=FALSE /USER=CEGID /PASSWORD=crypt(CEGID) /DATE=01/02/2000 /DOSSIER=Gens
OkEnvParams:=FALSE ; OkVerifStructure:=TRUE ;
OkUser:=FALSE ; OkPass:=FALSE ; OkDate:=FALSE ; OkDoss:=FALSE ;
for i:=1 to ParamCount do
   BEGIN
   St:=ParamStr(i) ;
   Nom:=UpperCase(Trim(ReadTokenPipe(St,'='))) ;
   Value:=UpperCase(Trim(St)) ;
//Paramètres de connexion
   if Nom='/MAJSTRUCTURE' then BEGIN OkVerifStructure:=(Value<>'FALSE') ; END ;
   if Nom='/USER'     then BEGIN OkUser:=TRUE ; FUser.Text:=Value ; END ;
   if Nom='/PASSWORD' then BEGIN OkPass:=TRUE ; FPassWord.Text:=DecryptageSt(Value) ; END ;
   if Nom='/DATE'     then BEGIN OkDate:=TRUE ; FDateEntree.Text:=Value ; END ;
   if Nom='/DOSSIER'  then
      BEGIN
      j:=FDossier.Items.indexOf(Value) ;
      if (j>=0) and (j<FDossier.Items.Count) then
         BEGIN
         FDossier.ItemIndex:=j; OkDoss:=TRUE ;
//pas sur         if Value='###DP###' then V_PGI.RunFromLanceur:=TRUE ;
         END ;
      END ;
   END;
OkEnvParams:=(OkDoss) and (OkUser) and (OkPass) and (OkDate) ;
if Not OkEnvParams then
   BEGIN
   GB1.Visible:=True;
   i:=0 ;
   ChargeDossier(FDossier.Items,True,True) ;
   i:=GetSynRegKey('LastSoc',0,True) ;
   if (i>=0) and (i<FDossier.Items.Count) then FDossier.ItemIndex:=i else FDossier.ItemIndex:=0 ;
   FDossierClick(Nil) ;
   END ;
FirstEnvParams:=FALSE ;
end;

procedure TMenuManuel.ActiveFonction(Ok : Boolean) ;
Var l : Integer ;
    CC1    : TControl;
BEGIN
For l:=0 To GB1.ControlCount-1 Do
  BEGIN
  CC1:=GB1.Controls[l] ;
  If CC1 Is TToolbarButton97 Then
    BEGIN
    TToolbarButton97(CC1).Enabled:=Ok ;
    If (TToolbarButton97(CC1).Tag=1) And (SuperUser) Then TToolbarButton97(CC1).Visible:=TRUE ;
    END ;
  END ;
FDossier.Visible:=Not Ok ; FPassword.Visible:=Not Ok ;
FUser.Visible:=Not Ok ; FDateEntree.Visible:=Not Ok ;
TFDossier.Visible:=Not Ok ; TFPassword.Visible:=Not Ok ;
TFUser.Visible:=Not Ok ; TFDateEntree.Visible:=Not Ok ;
Image1.Visible:=Ok ; ImageS5.Visible:=Ok ;
END ;

procedure TMenuManuel.BConnectClick(Sender: TObject);
begin
  BDeconnect.Enabled := False;
  if (V_PGI.OkOuvert) then
  begin
    Fermeture;
    Exit ;
  end;
  V_PGI.Versiondev:=FALSE ;
  V_PGI.Synap:=FALSE ;
  V_PGI.DateEntree:=StrToDate(FDateEntree.Text) ;
  V_PGI.UserName:=FUser.Text ; V_PGI.PassWord:=CryptageSt(FPassWord.text) ;
  VH^.GrpMontantMin:=0 ;
  VH^.GrpMontantMax:=1000000 ;
  VH^.GereSousPlan:=True ;
  VH^.Mugler:=FALSE ;
  VH^.JalVTP:='OTP' ; VH^.JalATP:='OTP' ;
  OkEnvParams:=FALSE ; bConnect.Enabled:=FALSE ; bQuitter.Enabled:=FALSE ;
  // ------------------- MD ----------------------
  //# traitements PGIEnv
  if V_PGI_Env<>Nil then
  begin
    //# mono-entreprise ou multi depuis lanceur : pointe l'environnement de la soc choisie
    if (V_PGI_Env.ModeFonc='MONO') or ((V_PGI_Env.ModeFonc='MULTI') and (ctxlanceur in V_PGI.PGIContexte)) Then
      V_PGI_Env.SocCommune := FDossier.Items[FDossier.ItemIndex];
    //# mode multi-dossier : appli lancée sans passer par le lanceur...
    if (V_PGI_Env.ModeFonc='MULTI') and (Not V_PGI.RunFromLanceur) and (Not (ctxlanceur in V_PGI.PGIContexte)) then
   //manque le double choix : soc commune, dossier
    PGIInfo('Vous lancez une application en mode multi-dossier sans passer par le lanceur.', TitreHalley);
  end;
  // ------------------- FIN MD ----------------------
  if ConnecteHalley(FDossier.Items[FDossier.ItemIndex],True,@ChargeMagHalley,nil,nil,nil,FALSE) then
  begin
    V_PGI.OKOuvert:=TRUE ; V_PGI.CurrentAlias:=FDossier.Items[FDossier.ItemIndex] ;
    ChargeAccesConcept ;
    if V_PGI.Tablette then ChargerLesTables ;
    BDeconnect.Visible:=True ;
    BConnect.Visible:=FALSE ;
    BQuitter.Visible:=TRUE ;
    GB1.Enabled:= True;    
    if V_PGI.RunFromLanceur then BEGIN Delay(1000) ;  END ;
    if V_PGI.RunFromLanceur then
    begin
      //Visible := False;
      (* GP Pour CA : A vérifier ???
      EtatPCL;
      Close;
      *)
    end else
    begin
      BDeconnect.Enabled := True;
      ActiveFonction(True);
    end;
  end else
  BEGIN
    BConnect.Enabled:=TRUE ;
    If (DBSOC<>NIL) AND (DBSOC.Connected) then DBSOC.Connected:=FALSE ;
  end ;
  SourisNormale ;
  BConnect.Enabled:=Not V_PGI.RunFromLanceur ;
  BQuitter.Enabled:=TRUE ;
end ;

procedure TMenuManuel.BDeconnectClick(Sender: TObject);
begin
BDeconnect.Visible:=FALSE ;
ActiveFonction(FALSE) ;
GB1.Enabled:=False ;
Fermeture;
BConnect.Visible:=TRUE ;
end;

procedure TMenuManuel.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Timer.Enabled:=FALSE ;
if V_PGI.OKOuvert then
   BEGIN
   Logout ;
   DeconnecteHalley ;
   END ;
end;

procedure TMenuManuel.BquitterClick(Sender: TObject);
begin Close ; end;

procedure TMenuManuel.FormKeyPress(Sender: TObject; var Key: Char);
begin
if Ord(Key) = 27 then bQuitterClick(nil) ;
if (not V_PGI.OKOuvert) and  (Ord(Key) = 13) then bConnectClick(nil) ;
if not V_PGI.OKOuvert then exit ;
end;

procedure TMenuManuel.BEtatsLibresClick(Sender: TObject);
begin
EtatPCL(esBil) ;
end;

procedure TMenuManuel.FormCreate(Sender: TObject);
begin
FirstEnvParams:=TRUE ;
if V_PGI.CegidBureau then
   BEGIN
   PGIAppAlone:=True ;
   CreatePGIApp ;
   END ;
end;

procedure TMenuManuel.FDossierClick(Sender: TObject);
Var IniFile : TIniFile ;
BEGIN
if FDossier.ItemIndex<0 then exit ;
IniFile :=TIniFile.Create(HALSOCINI);
FUser.Text:='' ;
FUser.Text:=FUser.Text+IniFile.ReadString(FDossier.Items[FDossier.ItemIndex],'LastUser','');
IniFile.Free;
SaveSynRegKey('LastSoc',FDossier.ItemIndex,True) ;
end;

procedure TMenuManuel.Fermeture;
var    i : integer ;
begin
  FDateEntree.Text:=DateToStr(Date) ;
  i:=0 ;
  //# lecture classique halsocini, VireDP=False la 1ère fois
  ChargeDossier(FDossier.Items,TRUE,Not FirstEnvParams) ;
  i:=GetSynRegKey('LastSoc',0,True) ;
  if (i>=0) and (i<FDossier.Items.Count) then FDossier.ItemIndex:=i else FDossier.ItemIndex:=0 ;
  FDossierClick(Nil) ;
  if V_PGI.OKOuvert then // Fermer
  begin
    LogOut ;
    DeconnecteHalley ;
    V_PGI.OKOuvert:=FALSE ; V_PGI.LanguePrinc:='FRA' ; V_PGI.LanguePerso:='FRA' ;
  end ;
  BDeconnect.Visible:=FALSE ;
  BConnect.Visible:=TRUE ; BConnect.Enabled:=TRUE ;
  bQuitter.Visible:=TRUE    ; bQuitter.Enabled:=TRUE ;
  FocusControle(FUSer) ;
  If FDossier.CanFocus And (FDossier.Items.Count>1) Then FDossier.SetFocus ;
end;

procedure TMenuManuel.TimerTimer(Sender: TObject);
Var ok : boolean ;
begin
  if Timer.Interval<2000 then
  BEGIN
    Visible:=True ;
    Timer.Interval:=Timer.Interval*10 ;
    Application.ProcessMessages ;
    if OkEnvParams then  BConnectClick(nil) ;
  END else if V_PGI.OKOuvert then
  BEGIN
   Timer.Interval:=V_PGI.TimerMessage*1000 ;
  END ;
end;

procedure TMenuManuel.ToolbarButton971Click(Sender: TObject);
begin
EtatPCL(esCR) ;
end;

procedure TMenuManuel.ToolbarButton972Click(Sender: TObject);
begin
EtatPCL(esSIG) ;
end;

initialization
{Titres et série du programme}
Apalatys:= 'CEGID' ;
HalSocIni:='CEGIDPGI.ini' ;
NomHalley:= 'Comptabilité S5' ;
TitreHalley := 'CEGID Comptabilité S5' ;
Copyright   := '© Copyright CEGID' ;
V_PGI.NumVersion:='3.80' ;
V_PGI.NumBuild:='525' ;
V_PGI.DateVersion:=EncodeDate(2001,02,09) ;
V_PGI.CegidApalatys := False;
V_PGI.CegidBureau := True;
V_PGI.Sav := True;
V_PGI.Decla100 := True;
V_PGI.MultiUserLogin := True;
V_PGI.StandardSurDP := True;
V_PGI.MajPredefini := True;
V_PGI.VersionReseau := true;
V_PGI.BlockMAJStruct:=True ;
V_PGI.EuroCertifiee:=False ;
V_PGI.LaSerie:=S5 ;
ChargeXuelib ;


end.
