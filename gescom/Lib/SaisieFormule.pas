unit SaisieFormule;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, Hctrls;

type
  TFSaisieFormule = class(TForm)
    ListeTable: TListBox;
    ListeFichier: TListBox;
    Edit1: TEdit;
    BitBtn1: THBitBtn;
    BitBtn2: THBitBtn;
    BitBtn3: THBitBtn;
    BitBtn4: THBitBtn;
    BitBtn5: THBitBtn;
    BitBtn6: THBitBtn;
    BitBtn7: THBitBtn;
    BitBtn8: THBitBtn;
    BitBtn9: THBitBtn;
    BitBtn10: THBitBtn;
    BitBtn11: THBitBtn;
    BitBtn12: THBitBtn;
    BitBtn13: THBitBtn;
    BitBtn14: THBitBtn;
    BitBtn15: THBitBtn;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    FctsDisp: THValComboBox;
    HLabel1: THLabel;
    ListeTableCode: TListBox;
    ListeFichierCode: TListBox;
    procedure BitBtn(Sender: TObject);
    procedure ListeTableDblClick(Sender: TObject);
    procedure ListeFichierDblClick(Sender: TObject);
    procedure FctsDispChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

var
  FSaisieFormule: TFSaisieFormule;

implementation

{$R *.DFM}

procedure TFSaisieFormule.BitBtn(Sender: TObject);
var
    st_trav1 : string;
    ch_trav1 : Char;

begin
    if Sender = BitBtn1 then ch_trav1 := '1';
    if Sender = BitBtn2 then ch_trav1 := '2';
    if Sender = BitBtn3 then ch_trav1 := '3';
    if Sender = BitBtn4 then ch_trav1 := '4';
    if Sender = BitBtn5 then ch_trav1 := '5';
    if Sender = BitBtn6 then ch_trav1 := '6';
    if Sender = BitBtn7 then ch_trav1 := '7';
    if Sender = BitBtn8 then ch_trav1 := '8';
    if Sender = BitBtn9 then ch_trav1 := '9';
    if Sender = BitBtn10 then ch_trav1 := '0';
    if Sender = BitBtn11 then ch_trav1 := '+';
    if Sender = BitBtn12 then ch_trav1 := '-';
    if Sender = BitBtn13 then ch_trav1 := '*';
    if Sender = BitBtn14 then ch_trav1 := '/';
    if Sender = BitBtn15 then ch_trav1 := '.';
    st_trav1 := Copy(Edit1.Text, 0, Edit1.SelStart) + ch_trav1;
    Edit1.Text := st_trav1 + Copy(Edit1.Text, Edit1.SelStart + 1, 255);
    Edit1.SetFocus;
    Edit1.SelStart := Length(st_trav1);
    Edit1.SelLength := 0;
end;

procedure TFSaisieFormule.ListeTableDblClick(Sender: TObject);
var
    st_trav1 : string;

begin
//    Edit1.Text := Edit1.Text + ListeTable.Items.Strings[ListeTable.ItemIndex];
    st_trav1 := Copy(Edit1.Text, 0, Edit1.SelStart) + ListeTableCode.Items.Strings[ListeTable.ItemIndex];
    Edit1.Text := st_trav1 + Copy(Edit1.Text, Edit1.SelStart + 1, 255);
    Edit1.SetFocus;
    Edit1.SelStart := Length(st_trav1);
    Edit1.SelLength := 0;
end;

procedure TFSaisieFormule.ListeFichierDblClick(Sender: TObject);
var
    st_trav1 : string;

begin
//    Edit1.Text := Edit1.Text + ListeFichier.Items.Strings[ListeFichier.ItemIndex];
    st_trav1 := Copy(Edit1.Text, 0, Edit1.SelStart) + ListeFichierCode.Items.Strings[ListeFichier.ItemIndex];
    Edit1.Text := st_trav1 + Copy(Edit1.Text, Edit1.SelStart + 1, 255);
    Edit1.SetFocus;
    Edit1.SelStart := Length(st_trav1);
    Edit1.SelLength := 0;
end;

procedure TFSaisieFormule.FctsDispChange(Sender: TObject);
begin
Edit1.Text := Edit1.Text + FctsDisp.Value;
if FctsDisp.Value = 'COMPTEUR' then Edit1.Hint := 'COMPTEUR';
if FctsDisp.Value = 'REPLACECAR' then
    begin
    Edit1.Text := Edit1.Text + '(';
    Edit1.Hint := 'REPLACECAR( Champ(s) à traiter, Chaine à remplacer, Chaine de remplacement )';
    end;
if FctsDisp.Value = 'CORRESPOND' then
    begin
    Edit1.Text := Edit1.Text + '(';
    Edit1.Hint := 'CORRESPOND( Champ(s) à traiter )';
    end;
if FctsDisp.Value = 'MAJUSCULE' then
    begin
    Edit1.Text := Edit1.Text + '(';
    Edit1.Hint := 'MAJUSCULE( Champ(s) à traiter )';
    end;
if FctsDisp.Value = 'TABLE' then
    begin
    Edit1.Text := Edit1.Text + '(';
    Edit1.Hint := 'TABLE( Table à lire )';
    end;
if FctsDisp.Value = 'SOUSCHAINE' then
    begin
    Edit1.Text := Edit1.Text + '(';
    Edit1.Hint := 'SOUSCHAINE( Champ à traiter, Début, Longueur )';
    end;
if FctsDisp.Value = 'SPACE' then
    begin
    Edit1.Text := Edit1.Text + '(';
    Edit1.Hint := 'SPACE( Nombre d''espaces à insérer )';
    end;
if FctsDisp.Value = 'INTSTR' then
    begin
    Edit1.Text := Edit1.Text + '(';
    Edit1.Hint := 'INTSTR( Champ à convertir )';
    end;
if FctsDisp.Value = 'STRINT' then
    begin
    Edit1.Text := Edit1.Text + '(';
    Edit1.Hint := 'STRINT( Champ à convertir )';
    end;
if FctsDisp.Value = 'DATE' then
    begin
    Edit1.Text := Edit1.Text + '(';
    Edit1.Hint := 'DATE( Valeur de la date )';
    end;
Edit1.SetFocus;
Edit1.SelStart := Length(Edit1.Text);
Edit1.SelLength := 0;
end;

procedure TFSaisieFormule.FormClose(Sender: TObject; var Action: TCloseAction);
var
    st1 : string;
    i_ind1, i_ind2 : integer;

begin
if Pos('REPLACECAR', Edit1.Text) <> 0 then
    begin
    if Edit1.Text[Length(Edit1.Text)] <> ')' then Edit1.Text := Edit1.Text + ')';
    st1 := Copy(Edit1.Text, Pos('(', Edit1.Text) + 1, Length(Edit1.Text) - (Pos('(', Edit1.Text) + 1));
    i_ind2 := 0;
    for i_ind1 := 1 to Length(st1) - 1 do
        if st1[i_ind1] = ',' then Inc(i_ind2);
    if i_ind2 <> 2 then Action := caNone;
    end;
if Pos('CORRESPOND', Edit1.Text) <> 0 then
    begin
    st1 := Trim(Copy(Edit1.Text, Pos('(', Edit1.Text) + 1, Length(Edit1.Text) - (Pos('(', Edit1.Text) + 1)));
    if st1 = '' then Action := caNone;
    if Edit1.Text[Length(Edit1.Text)] <> ')' then Edit1.Text := Edit1.Text + ')';
    end;
if Pos('MAJUSCULE', Edit1.Text) <> 0 then
    begin
    st1 := Trim(Copy(Edit1.Text, Pos('(', Edit1.Text) + 1, Length(Edit1.Text) - (Pos('(', Edit1.Text) + 1)));
    if st1 = '' then Action := caNone;
    if Edit1.Text[Length(Edit1.Text)] <> ')' then Edit1.Text := Edit1.Text + ')';
    end;
if Pos('TABLE', Edit1.Text) <> 0 then
    begin
    st1 := Trim(Copy(Edit1.Text, Pos('(', Edit1.Text) + 1, Length(Edit1.Text) - (Pos('(', Edit1.Text) + 1)));
    if st1 = '' then Action := caNone;
    if Edit1.Text[Length(Edit1.Text)] <> ')' then Edit1.Text := Edit1.Text + ')';
    end;
if Pos('DATE', Edit1.Text) <> 0 then
    begin
    st1 := Trim(Copy(Edit1.Text, Pos('(', Edit1.Text) + 1, Length(Edit1.Text) - (Pos('(', Edit1.Text) + 1)));
    if st1 = '' then Action := caNone;
    if Edit1.Text[Length(Edit1.Text)] <> ')' then Edit1.Text := Edit1.Text + ')';
    end;
end;

end.
