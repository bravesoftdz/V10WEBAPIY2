unit SaisieAssTablette;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HTB97, StdCtrls, ExtCtrls, HPanel;

type
  TFSelAssociation = class(TForm)
    HPanel2: THPanel;
    Dock971: TDock97;
    ListeTable: TListBox;
    ToolWindow971: TToolWindow97;
    BAide: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BValider: TToolbarButton97;
    ListeTableNam: TListBox;
    Panel1: TPanel;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    procedure BValiderClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure RadioButton1Click(Sender: TObject);
    procedure RadioButton2Click(Sender: TObject);
    procedure ListeTableNamClick(Sender: TObject);
    procedure ListeTableClick(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
    ValRetour : string;
  end;

var
  FSelAssociation: TFSelAssociation;

implementation

{$R *.DFM}

procedure TFSelAssociation.BValiderClick(Sender: TObject);
begin
if ListeTable.ItemIndex >= 0 then ValRetour := ListeTable.Items.Strings [ListeTable.ItemIndex];
close;
end;

procedure TFSelAssociation.FormShow(Sender: TObject);
begin
ValRetour := '';
RadioButton1.Checked := true;
end;

procedure TFSelAssociation.BAbandonClick(Sender: TObject);
begin
close;
end;

procedure TFSelAssociation.RadioButton1Click(Sender: TObject);
begin
ListeTableNam.Visible := true;
ListeTable.Visible := false;
end;

procedure TFSelAssociation.RadioButton2Click(Sender: TObject);
begin
ListeTable.Visible := true;
ListeTableNam.Visible := false;
end;

procedure TFSelAssociation.ListeTableNamClick(Sender: TObject);
begin
ListeTable.ItemIndex := ListeTableNam.ItemIndex;
end;

procedure TFSelAssociation.ListeTableClick(Sender: TObject);
begin
ListeTableNam.ItemIndex := ListeTable.ItemIndex;
end;

end.
