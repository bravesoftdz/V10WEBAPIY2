unit BTSelectDate;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, NewCalendar,
  Vierge, HSysMenu, HTB97;

type
  TFselectDate = class(TFVierge)
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
  private
    { Déclarations privées }
    TheSelector : TmxCalendar;
    fDateDeb,FdateFin,FSelected : TDateTime;
  public
    { Déclarations publiques }
    property DateDebut : TdateTime read fDateDeb write fDateDeb;
    property DateFin : TdateTime read fDatefin write fDatefin;
    property Selected : TdateTime read fSelected;
  end;

var
  FselectDate: TFselectDate;

function SelectDateFromCalendar (DateDepart,DateFin : TdateTime) : TdateTime;

implementation

{$R *.DFM}

function SelectDateFromCalendar (DateDepart,DateFin : TdateTime) : TdateTime;
var FF : TFSelectDate;
begin
  result := 0;
  FF := TFselectDate.create (Application);
  TRY
    FF.DateDebut := DateDepart;
    FF.DateFin := DateFin;
    FF.ShowModal;
    result := FF.Selected;
  FINALLY
    FF.free;
  END;
end;

{ TFselectDate }

procedure TFselectDate.FormCreate(Sender: TObject);
begin
  inherited;
  TheSelector := TmxCalendar.create (self);
  TheSelector.Parent := self;
  TheSelector.Align := alClient;
  FSelected := 0;
end;

procedure TFselectDate.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  inherited;
  TheSelector.free;
end;

procedure TFselectDate.FormShow(Sender: TObject);
begin
  inherited;
  TheSelector.MinDate := DateDebut;
  TheSelector.MaxDate := DateFin;
  TheSelector.Invalidate;
  TheSelector.Repaint;
end;

procedure TFselectDate.BValiderClick(Sender: TObject);
begin
  inherited;
  fSelected := TheSelector.SelectionStart;
end;

end.
