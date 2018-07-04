{***********UNITE*************************************************
Auteur  ...... : Compta
Créé le ...... : 25/02/2003
Modifié le ... :   /  /
Description .. : Passage en eAGL
Mots clefs ... :
*****************************************************************}
unit ClotANA;

interface

uses
  Windows, Classes, Controls, Forms, StdCtrls, ExtCtrls, Hctrls,
  Mask,   // TMaskEdit
  Ent1,   // TExoDate, TSetFichierBase, TFichierBase, AxeTofb, jbJal, VH, ExoToDates
  HEnt1,  // String3, SyncrDefault, BeginTrans, CommitTrans, RollBack, NumSemaine, V_PGI, EnableControls
{$IFDEF EAGLCLIENT}
  uLanceProcess,
{$ELSE}
  DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
  SaisUtil,  // SetIncNum
  HStatus,   // InitMove, MoveCur, FinitMove
  HSysMenu,
  SoldeCpt,  // MajTotComptes
  HPanel,    // THPanel
  UiUtil,    // FindInsidePanel, InitInside
  uLibClotureAna, // Traitements
  SysUtils, //IntToStr
  ULibAnalytique, //RecherchePremDerAxeVentil
  UTOB,
  hmsgbox,
  HTB97;

Function ClotureComptableAna : Boolean ;

Type TLExo = Record
             Code : String ;
             Deb,Fin : TDateTime ;
             END ;

type
  TFCloAna = class(TForm)
    HMTrad  : THSystemMenu;
    HMess   : THMsgBox;
    Dock971: TDock97;
    HPB: TToolWindow97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    GBencours: TGroupBox;
    EnCours: TLabel;
    LPatience: TLabel;
    GroupBox2: TGroupBox;
    Cpt1: TLabel;
    Cpt2: TLabel;
    GroupBox1: TGroupBox;
    HCpt1: TLabel;
    Panel2: TPanel;
    Label6: TLabel;
    GBOuvre: TGroupBox;
    HLabel10: THLabel;
    Label7: TLabel;
    HLabel3: THLabel;
    HLabel13: THLabel;
    HLabel14: THLabel;
    Label1: TLabel;
    ExoFaux: TLabel;
    FExoNew: THValComboBox;
    FDateClo1: TMaskEdit;
    FDateClo2: TMaskEdit;
    FExoClo: THValComboBox;
    FPourODA: TCheckBox;
    FPourChaPro: TCheckBox;
    GroupBox3: TGroupBox;
    HLabel6: THLabel;
    HLabel8: THLabel;
    HLabel9: THLabel;
    FJalODA: THValComboBox;
    FRefODA: TEdit;
    FLibODA: TEdit;
    GroupBox4: TGroupBox;
    HLabel2: THLabel;
    HLabel1: THLabel;
    FAxeODA: THValComboBox;
    FTypeCloODA: THValComboBox;
    FAvecANODA: TCheckBox;
    GroupBox5: TGroupBox;
    HLabel4: THLabel;
    HLabel5: THLabel;
    FAxeChaPro: THValComboBox;
    FTypeCloChaPro: THValComboBox;
    FAvecANChaPro: TCheckBox;
    GroupBox6: TGroupBox;
    HLabel7: THLabel;
    HLabel11: THLabel;
    HLabel12: THLabel;
    FJalChaPro: THValComboBox;
    FRefChaPro: TEdit;
    FLibChaPro: TEdit;
    FDateNew1: TMaskEdit;
    FDateNew2: TMaskEdit;
    GBPatienceCWAS: TGroupBox;
    Label2: TLabel;
    Label3: TLabel;
    procedure BValiderClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure FPourODAClick(Sender: TObject);
    procedure FPourChaProClick(Sender: TObject);
    procedure FExoNewChange(Sender: TObject);
    procedure FExoCloChange(Sender: TObject);
    procedure FTypeCloChaProChange(Sender: TObject);
    procedure FTypeCloODAChange(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
  private
    { Déclarations privées }
    // objet qui va faire tout le boulot ;)
    ClotureProcess : TTraitementClotureAna ;

    OnSort : Boolean ;
    Lexo : Array[0..30] Of TLExo ;
    Function InitParam(Var F : TFiltreCloAna) : Integer ;
    Procedure ActiveChamp(OkOk : Boolean ; What : Integer) ;
{$IFDEF EAGLCLIENT}
    procedure ClotureProcessServer ;
{$ENDIF}
    procedure Cloture2Tiers;
  public
    { Déclarations publiques }
    Procedure Mess(St1,St2,St3 : String) ;
    Procedure Patience ( vBoEnable : Boolean ) ;
    Procedure EcranEnCours ;
  end;


implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  ULibExercice,
  {$ENDIF MODENT1}
  UtilPgi;  // _BlocageMonoPoste, _DeblocageMonoPoste

{$R *.DFM}
Function ClotureComptableAna : Boolean ;
var FClo: TFCloAna;
    OutProg : Boolean ;
    PP : THPanel ;
begin
Result:=FALSE ;
if Not _BlocageMonoPoste(True) then Exit ;
FClo:=TFCloAna.Create(Application) ; OutProg:=FALSE ;
FClo.OnSort:=OutProg ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FClo.ShowModal ;
    Finally
     OutProg:=FClo.OnSort ;
     FClo.free ;
     _DeblocageMonoPoste(True) ;
    End ;
   END else
   BEGIN
   InitInside(FClo,PP) ;
   FClo.Show ;
   END ;
If OutProg  Then Result:=TRUE ;
Screen.Cursor:=SyncrDefault ;
end ;

Procedure TFCloAna.Mess(St1,St2,St3 : String) ;
begin
HCpt1.Caption:=St1 ; Cpt1.Caption:=St2 ; Cpt2.Caption:=St3 ;
Application.ProcessMessages ;
end ;

Function TFCloAna.InitParam(Var F : TFiltreCloAna) : Integer ;
Var Q : TQuery ;
    DD : TDateTime ;
BEGIN
Result:=0 ;
Fillchar(F,SizeOf(F),#0) ;
Q:=OpenSQL('SELECT * FROM EXERCICE WHERE EX_EXERCICE="'+FExoClo.Value+'" ',TRUE) ;
If Not Q.Eof Then
  BEGIN
  F.GeneODA.CloExo.Code:=Q.FindField('EX_EXERCICE').AsString ;
  F.GeneODA.CloExo.Deb:=Q.FindField('EX_DATEDEBUT').AsDateTime ;
  F.GeneODA.CloExo.Fin:=Q.FindField('EX_DATEFIN').AsDateTime ;
  F.GeneChaPro.CloExo.Code:=F.GeneODA.CloExo.Code ;
  F.GeneChaPro.CloExo.Deb:=F.GeneODA.CloExo.Deb ;
  F.GeneChaPro.CloExo.Fin:=F.GeneODA.CloExo.Fin ;
  END ;
Ferme(Q) ;
If F.GeneODA.CloExo.Code='' Then BEGIN Result:=1 ; Exit ; END ;
Q:=OpenSQL('SELECT * FROM EXERCICE ORDER BY EX_DATEDEBUT',TRUE) ;
While Not Q.Eof Do
  BEGIN
  DD:=Q.FindField('EX_DATEDEBUT').AsDateTime ;
  If DD>F.GeneODA.CloExo.Fin Then
    BEGIN
    F.GeneODA.NewExo.Code:=Q.FindField('EX_EXERCICE').AsString ;
    F.GeneODA.NewExo.Deb:=Q.FindField('EX_DATEDEBUT').AsDateTime ;
    F.GeneODA.NewExo.Fin:=Q.FindField('EX_DATEFIN').AsDateTime ;
    F.GeneChaPro.NewExo.Code:=F.GeneODA.NewExo.Code ;
    F.GeneChaPro.NewExo.Deb:=F.GeneODA.NewExo.Deb ;
    F.GeneChaPro.NewExo.Fin:=F.GeneODA.NewExo.Fin ;
    Break ;
    END ;
  Q.Next ;
  END ;
Ferme(Q) ;
If F.GeneODA.NewExo.Code='' Then BEGIN Result:=2 ; Exit ; END ;
If FPourODA.Checked Then With F.GeneODA Do
  BEGIN
  // FQ 10424
  Jal:=FJalODA.Value ; Ref:=FRefODA.Text ; Lib:=FLibODA.Text ;
  //SG6 07.03.05 mode croisaxe
  if not VH^.AnaCroisaxe then
    Axe:=FAxeODA.Value
  else
    Axe := 'A' + IntToStr(RecherchePremDerAxeVentil.premier_axe);

  AvecANO:=FAvecANODA.Checked ;
  TypeCloAna:=EnDetail ; If FTypeCloODA.ItemIndex=1 Then TypeCloAna:=EnSolde ;
  Q:=OpenSQL('Select J_COMPTEURNORMAL,J_AXE from JOURNAL Where J_JOURNAL="'+Jal+'"',True) ;
  If Not Q.Eof Then
    BEGIN
    Souche:=Q.Fields[0].AsString ;
    If Q.Fields[1].AsString<>Axe Then Result:=4 ;
    END Else Result:=5 ;
  Ferme(Q) ;
  END ;
If FPourChaPro.Checked Then With F.GeneChaPro Do
  BEGIN
  // FQ 10424
  Jal:=FJalChaPro.Value ; Ref:=FRefChaPro.Text ; Lib:=FLibChaPro.Text ;
  //SG6 07.03.05 mode croisaxe
  if not VH^.AnaCroisaxe then
    Axe:=FAxeChaPro.Value
  else
    Axe := 'A' + IntToStr(RecherchePremDerAxeVentil.premier_axe);

  AvecANO:=FAvecANChaPro.Checked ;
  TypeCloAna:=EnDetail ; If FTypeCloChaPro.ItemIndex=1 Then TypeCloAna:=EnSolde ;
  Q:=OpenSQL('Select J_COMPTEURNORMAL,J_AXE from JOURNAL Where J_JOURNAL="'+Jal+'"',True) ;
  If Not Q.Eof Then
    BEGIN
    Souche:=Q.Fields[0].AsString ;
    If Q.Fields[1].AsString<>Axe Then Result:=6 ;
    END Else Result:=7 ;
  Ferme(Q) ;
  END ;
If Result=0 Then
  If (Not FPourODA.Checked) And (Not FPourChaPro.Checked) Then BEGIN Result:=3 ; Exit ; END ;
END ;


procedure TFCloAna.BValiderClick(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
  ClotureProcessServer ;
{$ELSE}
  Cloture2Tiers ;
{$ENDIF}
end ;

procedure TFCloAna.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFCloAna.FormClose(Sender: TObject; var Action: TCloseAction);
begin

  ClotureProcess.Free ;

if Parent is THPanel then
   BEGIN
   _DeblocageMonoPoste(True) ;
   Action:=caFree ;
   END ;
end;

procedure TFCloAna.FormShow(Sender: TObject);
Var i              : Integer ;
    Q              : TQuery ;
begin
  //SG6 07.03.05 Analytique Croisaxe
  if VH^.Anacroisaxe then
  begin
    FAxeODA.Visible := False;
    HLabel2.Visible := False;
    HLabel4.Visible := False;
    FAxeChaPro.Visible := False;
  end;

  // Remplissage Liste des exo
  Fillchar(LExo,SizeOf(LExo),#0) ;
  i := 0 ;
  Q:=OpenSQL('Select * FROM EXERCICE ORDER BY EX_DATEDEBUT',TRUE) ;
  While Not Q.Eof Do
    BEGIN
    LExo[i].Code:=Q.FindField('EX_EXERCICE').AsString ;
    LExo[i].Deb:=Q.FindField('EX_DATEDEBUT').AsDateTime ;
    LExo[i].Fin:=Q.FindField('EX_DATEFIN').AsDateTime ;
    Inc(i) ;
    Q.Next ;
    END ;
  Ferme(Q) ;

  // Init interface
  FTypeCloODA.ItemIndex    := 1 ;
  FTypeCloChaPro.ItemIndex := 1 ;
  FPourODA.Checked         := FALSE ;
  FPourChaPro.Checked      := FALSE ;
  FExoClo.Value            := VH^.EnCours.Code ;
  FExoCloChange(nil);

  // Instanciation processus de cloture
  ClotureProcess := TTraitementClotureAna.Create( self, false ) ;
  GBPatienceCWAS.Visible := False ;
end;

Procedure TFCloAna.ActiveChamp(OkOk : Boolean ; What : Integer) ;
Var i,j : integer ;
    T   : TComponent ;
    GB : TGroupBox ;
BEGIN
for i:=0 to Self.ComponentCount-1 do
    BEGIN
    T:=Self.Components[i] ;
    If (T.Tag=What) And (T is TGroupBox) Then
      BEGIN
      GB:=TGroupBox(T) ; GB.Enabled:=OkOk ; 
      For j:=0 To GB.ControlCount-1 Do GB.Controls[j].Enabled:=OkOk ;
      END ;
    END ;
END ;

procedure TFCloAna.FPourODAClick(Sender: TObject);
begin
ActiveChamp(FPourODA.Checked,1) ;
FTypeCloODAChange(nil);
end;

procedure TFCloAna.FPourChaProClick(Sender: TObject);
begin
ActiveChamp(FPourChaPro.Checked,2) ;
FTypeCloChaProChange(nil);
end;

procedure TFCloAna.FExoNewChange(Sender: TObject);
begin
ExoToDates(FExoNew.Value,FDATENew1,FDATENew2) ;
end;

procedure TFCloAna.FExoCloChange(Sender: TObject);
Var i,j : Integer ;
begin
FExoNEw.Visible:=TRUE ; ExoFaux.Visible:=FALSE ;
ExoToDates(FExoClo.Value,FDATEClo1,FDATEClo2) ; j:=-1 ;
For i:=0 To 30 Do
  BEGIN
  If LExo[i].Code=FExoClo.Value Then j:=i+1 ;
  END ;
If (j<>-1) And (j<31) And (LExo[j].Code<>'') Then FExoNew.Value:=Lexo[j].Code Else
  BEGIN
  FExoNew.Visible:=FALSE ; ExoFaux.Visible:=TRUE ;
  END ;
end;

// FQ 10424
procedure TFCloAna.FTypeCloChaProChange(Sender: TObject);
var
  bEnabled : Boolean;
begin
  bEnabled := (FTypeCloChaPro.ItemIndex = 1);
  FRefChaPro.Enabled := bEnabled;
  FLibChaPro.Enabled := bEnabled;
  HLabel11.Enabled := bEnabled;
  HLabel12.Enabled := bEnabled;
end;

// FQ 10424
procedure TFCloAna.FTypeCloODAChange(Sender: TObject);
var
  bEnabled : Boolean;
begin
  bEnabled := (FTypeCloODA.ItemIndex = 1);
  FRefODA.Enabled := bEnabled;
  FLibODA.Enabled := bEnabled;
  HLabel8.Enabled := bEnabled;
  HLabel9.Enabled := bEnabled;
end;

procedure TFCloAna.Cloture2Tiers;
Var lParamClo : TFiltreCloAna ;
    errId     : Integer ;
begin
  // Vérification des paramètres
  errId := InitParam(lParamClo) ;
  if errId<>0 then
    begin
    HMess.Execute(errId+1,'','') ;
    Exit ;
    end ;

  // 2 confirmations
  if HMess.Execute(0,'','')<>mrYes then Exit ;
  if HMess.Execute(1,'','')<>mrYes Then Exit ;

  Screen.Cursor := SynCrDefault ;

  // Blocage interface pendant le traitement
  Patience(True) ;

  // Paramètres
  ClotureProcess.SetParamCloture( lParamClo ) ;

  // Traitement
  errId := ClotureProcess.ClotureAna ;

  // Déblocage de l'interface
  Patience(False) ;

  // Message final
  if errId = CLOANA_PASERREUR
    then HMess.Execute(15,'','')
    else
      begin
      HMess.Mess[14] := HMess.Mess[9] + HMess.Mess[10] + #13 + #13
                                      + HMess.Mess[11] + #13 + #13
                                      + HMess.Mess[12] + #13 + #13
                                      + HMess.Mess[13] + #13 + #13 ;
      HMess.Execute(14,'','') ;
      end ;

  Screen.Cursor := SynCrDefault ;

end;

{$IFDEF EAGLCLIENT}
procedure TFCloAna.ClotureProcessServer;
Var lParamClo : TFiltreCloAna ;
    errId     : Integer ;
    TobParam  : TOB ;
    TobResult : TOB ;
begin
  // Vérification des paramètres
  errId := InitParam(lParamClo) ;
  if errId<>0 then
    begin
    HMess.Execute(errId+1,'','') ;
    Exit ;
    end ;

  // 2 confirmations
  if HMess.Execute(0,'','')<>mrYes then Exit ;
  if HMess.Execute(1,'','')<>mrYes Then Exit ;

  Screen.Cursor := SynCrDefault ;

  // Blocage interface pendant le traitement
  Patience(True) ;

  // Préparation de la tob contenant les paramètres

  TobParam := ClotureProcess.CreerTobParamClo ( lParamClo ) ;

  // Traitement
  TobResult := LanceProcessServer('cgiCloture', 'clotureAna', 'aucun', TobParam, True ) ;

  // Récupération du résultat
  if TobResult.FieldExists('RESULT')
    then errID     := TobResult.GetValue('RESULT')
    else errID     := CLOANA_ERRPROCESSSERVER ;   // Pb avec le process server

  // Libération mémoire
  TobResult.Free ;
  TobParam.Free ;

  // Déblocage de l'interface
  Patience(False) ;

  // Message final
  if errId = CLOANA_ERRPROCESSSERVER
    then PGIBox('Attention : L''appel au processus serveur de clôture n''a pu aboutir. Veuillez vérifier votre installation.', Caption )
    else if errId <> CLOANA_PASERREUR then
           begin
           HMess.Mess[14] := HMess.Mess[9] + HMess.Mess[10] + #13 + #13
                                           + HMess.Mess[11] + #13 + #13
                                           + HMess.Mess[12] + #13 + #13
                                           + HMess.Mess[13] + #13 + #13 ;
           HMess.Execute(14,'','') ;
           end
         else HMess.Execute(15,'','') ;

  Screen.Cursor := SynCrDefault ;

end;

{$ENDIF}

procedure TFCloAna.Patience ( vBoEnable : Boolean ) ;
begin
  {$IFDEF EAGLCLIENT}
  GBOuvre.Visible        := not vBoEnable ;
  GBEnCours.Visible      := not vBoEnable ;
  GBPatienceCWAS.Visible := vBoEnable ;
  {$ELSE}
  if not vBoEnable then
    GBOuvre.Visible := True ;
  {$ENDIF EAGLCLIENT}
  EnableControls( Self, not vBoEnable ) ;
  Application.ProcessMessages;
end;

procedure TFCloAna.EcranEnCours;
begin
  GBOuvre.Visible   := FALSE ;
  GBEnCours.Visible := TRUE ;
end;


procedure TFCloAna.BFermeClick(Sender: TObject);
begin
  Close ;
  if IsInside(Self) then
    CloseInsidePanel(Self) ;
//if not IsInside(Self) then Close ;
//  Close;
end;

end.
