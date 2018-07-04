
unit PGGestionMP;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
{$IFNDEF EAGLCLIENT}
  Db,  DBGrids, HDB, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,
{$ELSE}
  MaineAGL,
{$ENDIF }
  HTB97, Grids,ExtCtrls, HPanel,  utob,
  StdCtrls, Hctrls, hent1, uiutil, Buttons, hmsgbox ;


// fonction de saisie de mot de passe
Function LanceControlMP (TypeAct : String) : Boolean;

type
  TPGGestionMP = class(TForm)
    HPanel1: THPanel;
    Dock971: TDock97;
    ToolWindow971: TToolWindow97;
    BVALIDER: TToolbarButton97;
    BFERMER: TToolbarButton97;
    BHELP: TToolbarButton97;
    Msg: THMsgBox;
    TMOTDEPASSE: TLabel;
    MOTDEPASSE: TEdit;
    NEWMOT: TEdit;
    TNEWMOT: TLabel;
    CNOUVMOT: TEdit;
    TCNOUVMOT: TLabel;
    procedure FormShow(Sender: TObject);
    procedure BVALIDERClick(Sender: TObject);
    procedure BHELPClick(Sender: TObject);
    procedure CNOUVMOTExit(Sender: TObject);
    procedure MOTDEPASSEExit(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private    { Déclarations privées }
    LeDroit : Boolean ;
    TypAct  : String ; // Action de la fiche : Saisie du mot de passe ou Modification/creation
    function ControleSaisie () : Integer ;
  end;

implementation

Uses  LicUtil,ParamSoc,EntPaie ;
{$R *.DFM}

Function LanceControlMP (TypeAct : String) : Boolean;
VAR
  X  : TPGGestionMP ;
  PP : THPanel ;
BEGIN
  X           := TPGGestionMP.Create (Application) ;
  X.TypAct    := TypeAct ;
  PP:=FindInsidePanel ;
  if PP=Nil then
   BEGIN
    try
     X.ShowModal ;
    finally
     X.Free ;
    end ;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;
END ;

PROCEDURE TPGGestionMP.FormShow(Sender: TObject);
var ST    : String;
    Q     : TQuery;
begin
  If TypAct = 'S' then
  begin
    MOTDEPASSE.visible := TRUE ;
    TMOTDEPASSE.visible := TRUE ;
    NEWMOT.visible := FALSE ;
    TNEWMOT.visible := FALSE ;
    CNOUVMOT.visible := FALSE ;
    TCNOUVMOT.visible := FALSE ;
  end
  else
  begin
    MOTDEPASSE.visible := FALSE ;
    TMOTDEPASSE.visible := FALSE ;
    NEWMOT.visible := TRUE ;
    TNEWMOT.visible := TRUE ;
    CNOUVMOT.visible := TRUE ;
    TCNOUVMOT.visible := TRUE ;
  end;
end;

procedure TPGGestionMP.BVALIDERClick(Sender: TObject);
begin
  ModalResult := mrNO ;
  ControleSaisie ();
  if LeDroit AND (TypAct <> 'C') then // Saisie
  begin
    SetParamSoc ('SO_PGTPCL', CryptageSt (NEWMOT.Text)) ;
    VH_Paie.PGMotDePasse := CryptageSt (NEWMOT.Text) ;
    ModalResult := mrYes ;
  end
  else
    if (TypAct = 'C') AND LeDroit then ModalResult := mrYes ;
end;

procedure TPGGestionMP.BHELPClick(Sender: TObject);
begin
  CallHelpTopic (Self);
end;

function TPGGestionMP.ControleSaisie: Integer;
begin
  if TypAct = 'C' then // Controle
  begin
  if NEWMOT.Text = DeCryptageSt (VH_Paie.PGMotDePasse) then
    LeDroit := TRUE
    else LeDroit := FALSE;
  end
  else
  begin
  if NEWMOT.Text = CNOUVMOT.Text then
    LeDroit := TRUE
    else LeDroit := FALSE;
  end;
end;

procedure TPGGestionMP.CNOUVMOTExit(Sender: TObject);
begin
  if NEWMOT.Text <> CNOUVMOT.Text then
  begin
    PgiBox ('Les nouveaux mots de passe sont différents,#13#10Entrez le nouveau mot de passe dans les 2 zônes', Caption);
    NEWMOT.SetFocus ;
    LeDroit := FALSE ;
  end
  else LeDroit := TRUE ;
end;

procedure TPGGestionMP.MOTDEPASSEExit(Sender: TObject);
begin
  BVALIDERClick (Sender) ;
end;

procedure TPGGestionMP.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  BVALIDERClick (Sender) ;
end;

end.
