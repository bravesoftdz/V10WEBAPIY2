unit CalcScor;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
{$IFDEF EAGLCLIENT}
  UTOB,
{$ELSE}
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
  StdCtrls, Buttons, ExtCtrls, HEnt1, HCtrls, Spin, Hcompte, HStatus,
  hmsgbox, Ent1, HSysMenu, ComCtrls, HTB97, HPanel, UiUtil ;

Procedure CalculScoring ;

type
  TFCalcScoring = class(TForm)
    PSelect: TGroupBox;
    EMin: TLabel;
    AUXI: THCpteEdit;
    Label1: TLabel;
    Label2: TLabel;
    AUXI_: THCpteEdit;
    Min: TSpinEdit;
    Patience: TPanel;
    EnCours: TLabel;
    LibAuxi: TLabel;
    MsgBox: THMsgBox;
    Anim: TAnimate;
    NbTiers: TLabel;
    Blabla: TLabel;
    HMTrad: THSystemMenu;
    HPB: TToolWindow97;
    Dock: TDock97;
    BStop: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    Bevel1: TBevel;
    procedure FormCreate(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BStopClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
    TTiers : TList ;
    Arreter : boolean ;
    procedure AppliqueScoring ;
    procedure InitComptes ;
  public
  end;

implementation

{$R *.DFM}

uses RapSuppr, UtilPgi ;

Procedure CalculScoring ;
var FCalcScoring:TFCalcScoring;
    PP : THPanel ;
begin

  if _Blocage(['nrCloture'],False,'nrBatch') then Exit ;

  FCalcScoring:=TFCalcScoring.Create(Application) ;
  PP := FindInsidePanel ;

  try
    if PP=Nil then
    begin
      try
        FCalcScoring.ShowModal ;
        finally
        FCalcScoring.Free ;
        end ;
    end
    else
    begin
      InitInside(FCalcScoring,PP) ;
      FCalcScoring.Show ;
    end ;
  finally
    _Bloqueur('nrBatch',False) ;
  end ;

  Screen.Cursor := SyncrDefault ;

END ;

procedure TFCalcScoring.FormCreate(Sender: TObject);
begin
  PopUpMenu := ADDMenuPop( PopUpMenu, '', '') ;
end;

procedure TFCalcScoring.BFermeClick(Sender: TObject);
begin
  Close ;
  if IsInside(Self) then
    CloseInsidePanel(Self) ;
end;

procedure TFCalcScoring.BValiderClick(Sender: TObject);
begin
  AppliqueScoring ;
end;

Procedure TFCalcScoring.AppliqueScoring ;
var lQTiers       : TQuery ;
    lQEcr         : TQuery ;
    lStSQL        : String ;
    NbCaption     : String ;
    Nb            : Integer ;
    ScoreRel      : SmallInt ;
    TotNiv        : LongInt ;
    X             : DelInfo ;
    OkCompte      : boolean ;
BEGIN
  Arreter := False ;

  // Test critères sur tiers
  if ((AUXI.Text='') or (AUXI_.Text='')) then
    begin
    Msgbox.Execute(0,'','') ;
    Exit ;
    end ;

  // Requête pour liste des clients lettrables
  lStSQL := 'SELECT T_SCORERELANCE, T_AUXILIAIRE, T_LIBELLE FROM TIERS' +
            ' WHERE T_AUXILIAIRE>="' + AUXI.Text + '" AND T_AUXILIAIRE<="' + AUXI_.Text + '"' +
              ' AND (T_NATUREAUXI="CLI" OR T_NATUREAUXI="AUD")' +
              ' AND T_LETTRABLE="X"' +
            ' ORDER BY T_AUXILIAIRE' ;
  lQTiers := OpenSQL( lStSQL, True ) ;

  // Aucun résultat
  if lQTiers.Eof then
    begin
    Msgbox.Execute(1,'','') ;
    Ferme(lQTiers) ;
    Exit ;
    end ;

  // Confirmation
  if Msgbox.Execute(3,'','')<>mrYes then
    begin
    Ferme(lQTiers) ;
    Exit ;
    end ;

  // Préparatin au traitement
  Application.ProcessMessages ;
  TTiers.Clear ;
  {$IFDEF EAGLCLIENT}
    InitMove( lQTiers.Detail.Count + 1, '') ;
  {$ELSE}
    InitMove( RecordsCount(lQTiers) + 1,'') ;
  {$ENDIF EAGLCLIENT}
  MoveCur(False) ;
  Blabla.Visible   := False ;
  Patience.Visible := True ;
  NbCaption        := Msgbox.Mess[5];
  OkCompte         := False ;
  Nb               := 0 ;

  // Parcours des tiers
  while not lQTiers.EOF do
    begin
    MoveCur(False) ;
    lStSQL := 'SELECT COUNT(E_NIVEAURELANCE) RELANCEQTE, SUM(E_NIVEAURELANCE) RELANCETOTAL' +
          ' FROM ECRITURE' +
          ' WHERE E_ECHE="X" AND E_NUMECHE>=0' +
            ' AND E_ETATLETTRAGE="TL"' +
            ' AND E_DEBIT>0' +
            ' AND E_AUXILIAIRE="' + lQTiers.FindField('T_AUXILIAIRE').AsString + '"' ;
    lQEcr := OpenSQL( lStSQL, True ) ;

    if ( ( lQEcr.FindField('RELANCEQTE').AsInteger < LongInt( Min.Value ) ) or
         ( lQEcr.FindField('RELANCEQTE').AsInteger = 0 ) ) then
       begin
       Ferme( lQEcr ) ;
       lQTiers.Next;
       Continue;
       end ;

    OkCompte := True ;
    Nb       := Nb + 1 ;

    NbTiers.Caption := NbCaption + IntToStr( Nb ) ;
    LibAUXI.Caption := lQTiers.FindField('T_LIBELLE').AsString ;

    Application.ProcessMessages ;
    if Arreter then
      begin
      Anim.Active := False ;
      Arreter     := False ;
      if Msgbox.Execute(4,'','') = mrYes then
        begin
        Ferme(lQEcr) ;
        Break ;
        end ;
      end ;

    TotNiv   := lQEcr.FindField('RELANCETOTAL').AsInteger + lQEcr.FindField('RELANCEQTE').AsInteger ;
    ScoreRel := ( ( TotNiv - 1 ) div lQEcr.FindField('RELANCEQTE').AsInteger ) + 1 ;

    Ferme( lQEcr ) ;

    // MAJ du tiers
    ExecuteSQL( 'UPDATE TIERS SET T_SCORERELANCE=' + IntToStr( ScoreRel ) +
                ' WHERE T_AUXILIAIRE="' + lQTiers.FindField('T_AUXILIAIRE').AsString  + '"' ) ;

    // Récup Tiers traité pour affichage infos
    X        := DelInfo.Create ;
    X.LeCod  := lQTiers.FindField('T_AUXILIAIRE').AsString ;
    X.LeLib  := lQTiers.FindField('T_LIBELLE').AsString ;
    X.LeMess := IntToStr( ScoreRel ) ;
    TTiers.Add(X) ;

    lQTiers.Next ;

    end ;

  Anim.Active := False ;
  FiniMove ;

  Ferme(lQTiers) ;

  if Not OkCompte
    then Msgbox.Execute(1,'','')
    else if MsgBox.Execute(2,'','') = mrYes then
           RapportDeSuppression( TTiers, 0 ) ;

  Patience.Visible := False ;
  BlaBla.Visible   := True ;
  Application.ProcessMessages ;

end ;

procedure TFCalcScoring.FormShow(Sender: TObject);
begin
  BlaBla.Visible := True ;
  Arreter        := False ;
  TTiers         := TList.Create ;
  InitComptes ;
end;

procedure TFCalcScoring.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  TTiers.Free ;
{
  if Parent is THPanel then
    Action := caFree ;
}
end;

procedure TFCalcScoring.BStopClick(Sender: TObject);
begin
  Arreter := True ;
end;

Procedure TFCalcScoring.InitComptes ;
var lQTiers : TQuery ;
begin
  lQTiers := OpenSQL('SELECT MIN(T_AUXILIAIRE) TIERSMIN, MAX(T_AUXILIAIRE) TIERSMAX FROM TIERS WHERE T_LETTRABLE="X"',True) ;
  if not lQTiers.Eof then
    begin
    AUXI.Text  := lQTiers.FindField('TIERSMIN').AsString ;
    AUXI_.Text := lQTiers.FindField('TIERSMAX').AsString ;
    end ;
  Ferme( lQTiers ) ;
END ;

procedure TFCalcScoring.BAideClick(Sender: TObject);
begin
  CallHelpTopic(Self) ;
end;

end.
