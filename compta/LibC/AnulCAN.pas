{***********UNITE*************************************************
Auteur  ...... : Compta
Créé le ...... : 04/03/2003
Modifié le ... :   /  /
Description .. : Passage en eAGL
Mots clefs ... :
*****************************************************************}
unit AnulCAN;

interface

uses
  Windows, Messages, Controls, Classes, Forms, StdCtrls,
  Dialogs,  // ShowMessage
{$IFDEF EAGLCLIENT}
  uLanceProcess,
{$ELSE}
  DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
  HPanel,  // THPanel
  ENT1,    // ExoToDates, VH
  HEnt1,   // SyncrDefault, EnableControls, BeginTrans, CommitTrans, RollBack
  UiUtil,  // FindInsidePanel, InitInside
  uLibClotureAna, // Traitements
  ULibAnalytique, //Recherche...
  UTOB,
  SysUtils,
  hmsgbox, Hctrls, HSysMenu, Mask, HTB97 ;

Procedure AnnuleClotureAnalytique ;

type
  TFDECLOANA = class(TForm)
    HPB       : TToolWindow97;
    BAide     : TToolbarButton97;
    BValider  : TToolbarButton97;
    BFerme    : TToolbarButton97;
    Confirmation: THMsgBox;
    HMTrad    : THSystemMenu;
    GroupBox1 : TGroupBox;
    Dock      : TDock97;
    FExoClo   : THValComboBox;
    FAxeODA   : THValComboBox;
    Label7    : TLabel;
    HLabel1   : THLabel;
    HLabel2   : THLabel;
    HLabel3   : THLabel;
    HLabel4   : THLabel;
    HLabel5   : THLabel;
    HLabel6   : THLabel;
    HLabel7   : THLabel;
    HLabel10  : THLabel;
    HLabel13  : THLabel;
    HPatienter: THLabel;
    FDateClo1 : TMaskEdit;
    FDateClo2 : TMaskEdit;
    FPourSaisie: TCheckBox;
    FPourODA   : TCheckBox;
    FPourChaPro: TCheckBox;
    GBFermeDEF : TGroupBox;
    procedure BValiderClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BAideClick(Sender: TObject);
    procedure FExoCloChange(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
  private
    // objet qui va faire tout le boulot ;)
    ClotureProcess : TTraitementClotureAna ;
{$IFDEF EAGLCLIENT}
    procedure AnnuleClotureProcessServer ;
{$ELSE}
    procedure AnnuleCloture2Tiers;
{$ENDIF EAGLCLIENT}
    procedure Patience( vBoEnable : Boolean ) ;
  public
  end;

implementation

{$R *.DFM}

Uses
  {$IFDEF MODENT1}
  ULibExercice,
  {$ENDIF MODENT1}
  ClotANA,   // tGenereANOAna, DetruitANOODA, MajTotClotureAna
  UtilPgi;   // _BlocageMonoPoste, _DeblocageMonoPoste

Procedure AnnuleClotureAnalytique ;
var FDeClo: TFDECLOANA;
    PP : THPanel ;
begin
if Not _BlocageMonoPoste(True) then Exit ;
FDEClo:=TFDECLOANA.Create(Application) ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FDEClo.ShowModal ;
    Finally
     FDEClo.free ;
     _DeblocageMonoPoste(True) ;
    End ;
   END else
   BEGIN
   InitInside(FDEClo,PP) ;
   FDEClo.Show ;
   END ;
Screen.Cursor:=SyncrDefault ;
end ;


procedure TFDECLOANA.BValiderClick(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
  AnnuleClotureProcessServer ;
{$ELSE}
  AnnuleCloture2Tiers ;
{$ENDIF}
end;

procedure TFDECLOANA.FormShow(Sender: TObject);
begin
  //SG6 07.03.05 Analytique croisaxe
  if VH^.AnaCroisaxe then
  begin
    HLabel5.Visible := False;
    FAxeODA.Visible := False;
  end;

  FExoClo.Value := VH^.EnCours.Code ;
  ExoToDates( VH^.EnCours.Code, FDateClo1, FDateClo2 ) ;
  FAXeODA.ItemIndex := 0 ;

  // Instanciation processus de cloture
  ClotureProcess := TTraitementClotureAna.Create( self, False ) ;

end;

procedure TFDECLOANA.FormClose(Sender: TObject; var Action: TCloseAction);
begin

  ClotureProcess.Free ;

  if Parent is THPanel then
    BEGIN
    _DeblocageMonoPoste(True) ;
    Action:=caFree ;
    END ;

end;

procedure TFDECLOANA.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFDECLOANA.FExoCloChange(Sender: TObject);
begin
ExoToDates(FExoClo.Value,FDATEClo1,FDATEClo2) ;
end;

{$IFDEF EAGLCLIENT}
procedure TFDECLOANA.AnnuleClotureProcessServer;
var errId     : Integer ;
    OkODA     : Boolean ;
    OkChaPro  : Boolean ;
    OkSaisie  : Boolean ;
    FG        : tGenereANOAna ;
    TobParam  : TOB ;
    TobResult : TOB ;
begin

  // Double confirmation
  errId:=Confirmation.Execute(0,'','') ;
  If errId<>mrYes Then Exit ;
  errId:=Confirmation.Execute(1,'','') ;
  If errId<>mrYes Then Exit ;

  // Blocage interface
  Patience( True ) ;

  // Préparation de la tob contenant les paramètres
  OkODA    := FPourODA.Checked ;
  OkChaPro := FPourChaPro.Checked ;
  OkSaisie := FPourSaisie.Checked ;
  Fillchar( FG, SizeOf(FG) , #0) ;
  FG.NewExo.Code := FExoClo.Value ;
  if not VH^.AnaCroisaxe then
    FG.Axe         := FAxeODA.Value
  else
    FG.Axe         := 'A' + IntToStr(RecherchePremDerAxeVentil.premier_axe);


  TobParam := ClotureProcess.CreerTobParamAnnulClo ( FG , OkODA, OkChaPro, OkSaisie ) ;

  // Traitement
  TobResult := LanceProcessServer('cgiCloture', 'declotureAna', 'aucun', TobParam, True ) ;

  // Récupération du résultat
  if TobResult.FieldExists('RESULT')
    then errID     := TobResult.GetValue('RESULT')
    else errID     := CLOANA_ERRPROCESSSERVER ;   // Pb avec le process server

  // Libération mémoire
  TobResult.Free ;
  TobParam.Free ;

  // Message final
  if errId = CLOANA_ERRPROCESSSERVER
    then PGIBox('Attention : L''appel au processus serveur de clôture n''a pu aboutir. Veuillez vérifier votre installation.', Caption )
    else if errId <> CLOANA_PASERREUR
           then ShowMessage(Confirmation.Mess[3])
           else Confirmation.Execute(2,'','') ;

  // Déblocage interface
  Patience( False ) ;
  BValider.Visible := FALSE ;

end;

{$ELSE}

procedure TFDECLOANA.AnnuleCloture2Tiers;
var errId    : Integer ;
    OkODA    : Boolean ;
    OkChaPro : Boolean ;
    OkSaisie : Boolean ;
    FG       : tGenereANOAna ;
begin


  OkODA    := FPourODA.Checked ;
  OkChaPro := FPourChaPro.Checked ;
  OkSaisie := FPourSaisie.Checked ;

  Fillchar( FG, SizeOf(FG) , #0) ;
  FG.NewExo.Code := FExoClo.Value ;
  if not VH^.AnaCroisaxe then
    FG.Axe         := FAxeODA.Value
  else
    FG.Axe         := 'A' + IntToStr(RecherchePremDerAxeVentil.premier_axe);


  // Double confirmation
  errId:=Confirmation.Execute(0,'','') ;
  If errId<>mrYes Then Exit ;
  errId:=Confirmation.Execute(1,'','') ;
  If errId<>mrYes Then Exit ;

  // Blocage interface
  Patience( True ) ;

  // Traitement
  ClotureProcess.SetParamAnnulClo ( FG, OkODA, OkChaPro, OkSaisie ) ;
  errId := ClotureProcess.AnnuleClotureAna ;

  // Message final
  if errId <> CLOANA_PASERREUR
    then ShowMessage(Confirmation.Mess[3])
    else Confirmation.Execute(2,'','') ;

  // Déblocage interface
  Patience( False ) ;
  BValider.Visible := FALSE ;

end;
{$ENDIF EAGLCLIENT}

procedure TFDECLOANA.Patience(vBoEnable: Boolean);
begin
  HPatienter.Visible:= vBoEnable ;
  EnableControls(Self, not vBoEnable) ;
  Application.ProcessMessages ;
end;

procedure TFDECLOANA.BFermeClick(Sender: TObject);
begin
  Close ;
  if IsInside(Self) then
    CloseInsidePanel(Self) ;
end;

end.
