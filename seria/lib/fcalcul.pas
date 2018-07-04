unit fcalcul;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, COmObj,CLKBtpLib_TLB, Mask, ComCtrls;

type
  Tcalcul = class(TForm)
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    resultat: TEdit;
    Label4: TLabel;
    Calcul: TButton;
    LIDsERv: TLabel;
    IdServeur: TMaskEdit;
    Cproduit: TMaskEdit;
    Cversion: TMaskEdit;
    CNbr: TEdit;
    GTempo: TGroupBox;
    CBTEMPO: TCheckBox;
    DTFin: TDateTimePicker;
    ldtFIn: TLabel;
    procedure CalculClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure CBTEMPOClick(Sender: TObject);
  private
    { Déclarations privées }
    WinSeria : TBuilder;
  public
    { Déclarations publiques }
  end;

var
  calcul: Tcalcul;

implementation

uses StrUtils, DateUtils;

{$R *.dfm}

procedure Tcalcul.CalculClick(Sender: TObject);
begin
	if (Cproduit.Text='') or (CVersion.text='') or (Cnbr.text='') then
  begin
  	Application.MessageBox('Merci de renseigner les champs pour le calcul','Erreur');
    exit;
  end;
  if not CBTEMPO.Checked then
  begin
  	Resultat.text := WinSeria.GetKey ( StrToInt(IdServeur.text)+21531,Widestring(CProduit.text+Cversion.text),StrToInt(Cnbr.text));
  end else
  begin
  	If (MonthOf (DTFin.Date) = MonthOf(Now)) then
    begin
      Application.MessageBox('La fin de validité est sur le mois courant','Erreur');
      exit;
    end;
  	Resultat.text := WinSeria.GetTempKey ( StrToInt(IdServeur.text)+21531,Widestring(CProduit.text+Cversion.text),StrToInt(Cnbr.text),MonthOf(DTFin.date),YearOf(DtFin.date));
  end;
end;

procedure Tcalcul.FormCreate(Sender: TObject);
begin
	WinSeria := TBuilDer.create(application);
end;

procedure Tcalcul.FormClose(Sender: TObject; var Action: TCloseAction);
begin
	if assigned(WinSeria) then WinSeria.free;
end;

procedure Tcalcul.CBTEMPOClick(Sender: TObject);
begin
	ldtFIn.visible := Tcheckbox(Sender).Checked;
	DTFin .visible := Tcheckbox(Sender).Checked;
end;

end.
