unit RecupTempoParamCli;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls,Dicobtp, hmsgbox, HSysMenu;

type
  TFParamCli = class(TForm)
    Label5: TLabel;
    Label4: TLabel;
    GroupBoxStatClient: TGroupBox;
    Label3: TLabel;
    NbStatCli: TLabel;
    Label7: TLabel;
    CheckStat1Cli: TCheckBox;
    CheckStat2Cli: TCheckBox;
    CheckStat3Cli: TCheckBox;
    CheckStat4Cli: TCheckBox;
    ComboStat1Cli: TComboBox;
    ComboStat2Cli: TComboBox;
    ComboStat3Cli: TComboBox;
    ComboStat4cli: TComboBox;
    GroupBoxGrilleTech: TGroupBox;
    Labelconseil: TLabel;
    LabelTableLibre3: TLabel;
    FermerParamFicBase: TButton;
    LabelTableLibre4: TLabel;
    LabelTableLibre5: TLabel;
    ComboTableLibre3: TComboBox;
    ComboTableLibre4: TComboBox;
    LabelTableLibre8: TLabel;
    ComboTableLibre5: TComboBox;
    LabelTableLibre9: TLabel;
    LabelTableLibre10: TLabel;
    ComboTableLibre8: TComboBox;
    ComboTableLibre9: TComboBox;
    ComboTableLibre10: TComboBox;
    EditCollectifClient: TEdit;
    LabelcollectifClient: TLabel;
    Labelcomptatiers: TLabel;
    ComboFamComptaTiers: TComboBox;
    Label6: TLabel;
    ComboTableLibre2: TComboBox;
    LabelTableLibre7: TLabel;
    ComboTableLibre7: TComboBox;
    LabelTableLibre1: TLabel;
    ComboTableLibre1: TComboBox;
    LabelTableLibre6: TLabel;
    ComboTableLibre6: TComboBox;
    HMtrad: THSystemMenu;
    ConcatAbv: TCheckBox;
    procedure FermerParamFicBaseClick(Sender: TObject);
    procedure CheckStat1CliClick(Sender: TObject);
    procedure CheckStat2CliClick(Sender: TObject);
    procedure CheckStat3CliClick(Sender: TObject);
    procedure CheckStat4CliClick(Sender: TObject);
    procedure ComboStat1CliClick(Sender: TObject);
    procedure ComboStat2CliClick(Sender: TObject);
    procedure ComboStat3CliClick(Sender: TObject);
    procedure ComboStat4cliClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure ComboTableLibre1Click(Sender: TObject);
    procedure ComboTableLibre2Click(Sender: TObject);
    procedure ComboTableLibre3Click(Sender: TObject);
    procedure ComboTableLibre4Click(Sender: TObject);
    procedure ComboTableLibre5Click(Sender: TObject);
    procedure ComboTableLibre6Click(Sender: TObject);
    procedure ComboTableLibre7Click(Sender: TObject);
    procedure ComboTableLibre8Click(Sender: TObject);
    procedure ComboTableLibre9Click(Sender: TObject);
    procedure ComboTableLibre10Click(Sender: TObject);
  private
    { Déclarations privées }
    function TestStat : Boolean;
    procedure TestStatUtilise ;
    function TestTableLibre : Boolean;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

procedure TFParamCli.FermerParamFicBaseClick(Sender: TObject);
begin
if (TestStat = False) then TWinControl(FermerParamFicBase).SetFocus
    else if (TestTableLibre = False) then TWinControl(FermerParamFicBase).SetFocus
        else Self.Close;
end;

procedure TFParamCli.CheckStat1CliClick(Sender: TObject);
begin
if (CheckStat1Cli.Checked = False) then ComboStat1Cli.Enabled := False
Else ComboStat1Cli.Enabled := True;
end;


procedure TFParamCli.CheckStat2CliClick(Sender: TObject);
begin
if (CheckStat2Cli.Checked = False) then ComboStat2Cli.Enabled := False
Else ComboStat2Cli.Enabled := True;
end;

procedure TFParamCli.CheckStat3CliClick(Sender: TObject);
begin
if (CheckStat3Cli.Checked = False) then ComboStat3Cli.Enabled := False
Else ComboStat3Cli.Enabled := True;
end;

procedure TFParamCli.CheckStat4CliClick(Sender: TObject);
begin
if (CheckStat4Cli.Checked = False) then ComboStat4Cli.Enabled := False
Else ComboStat4Cli.Enabled := True;
end;

procedure TFParamCli.ComboStat1CliClick(Sender: TObject);
begin
TestStatUtilise;
if (TestStat = False) then TWinControl(ComboStat1Cli).SetFocus;
end;

procedure TFParamCli.TestStatUtilise ;
Var NbStat,i: Integer;
Begin
NbStat := 0;
   // on reagrde si table libre tiers utilisée pour les STAT tempo/scot
   // si oui, on grise les tabmes libres correspondantes dans l'autre
   // partie d'écran.
If (ComboStat1Cli.ItemIndex = 0) Then Inc(NbStat);
If (ComboStat2Cli.ItemIndex = 0) Then Inc(NbStat);
If (ComboStat3Cli.ItemIndex = 0) Then Inc(NbStat);
If (ComboStat4Cli.ItemIndex = 0) Then Inc(NbStat);
ComboTableLibre1.Enabled := True; ComboTableLibre2.Enabled := True;
ComboTableLibre3.Enabled := True; ComboTableLibre4.Enabled := True;
For i:= 1 To NbStat do
    Begin
    Case i of
        1: Begin ComboTableLibre1.Enabled := False; ComboTableLibre1.ItemIndex := 0 End;
        2: Begin ComboTableLibre2.Enabled := False; ComboTableLibre2.ItemIndex := 0 End;
        3: Begin ComboTableLibre3.Enabled := False; ComboTableLibre3.ItemIndex := 0 End;
        4: Begin ComboTableLibre4.Enabled := False; ComboTableLibre4.ItemIndex := 0 End;
        End;
    End;
End;

function TFParamCli.TestStat : Boolean;
Var app, sect, Res1, Res2, Res3 : Integer;
Begin
     // on reagrde aussi si le choix apporteur et secteur n'a pas
     // été fait 2 fois
result:=TRUE;
app:=0;
sect :=0;
Res1 :=0;
Res2 :=0;
Res3:=0;
{ plus de gestion apporteur au 16/11/00 on deacle tout
If (ComboStat1Cli.ItemIndex = 1) Then Inc(app);
If (ComboStat2Cli.ItemIndex = 1) Then Inc(app);
If (ComboStat3Cli.ItemIndex = 1) Then Inc(app);
If (ComboStat4Cli.ItemIndex = 1) Then Inc(app);}
If (ComboStat1Cli.ItemIndex = 1) Then Inc(sect);
If (ComboStat2Cli.ItemIndex = 1) Then Inc(sect);
If (ComboStat3Cli.ItemIndex = 1) Then Inc(sect);
If (ComboStat4Cli.ItemIndex = 1) Then Inc(sect);
If (ComboStat1Cli.ItemIndex =2) Then Inc(Res1);
If (ComboStat2Cli.ItemIndex =2) Then Inc(Res1);
If (ComboStat3Cli.ItemIndex =2) Then Inc(Res1);
If (ComboStat4Cli.ItemIndex =2) Then Inc(Res1);
If (ComboStat1Cli.ItemIndex =3) Then Inc(Res2);
If (ComboStat2Cli.ItemIndex =3) Then Inc(Res2);
If (ComboStat3Cli.ItemIndex =3) Then Inc(Res2);
If (ComboStat4Cli.ItemIndex =3) Then Inc(Res2);
If (ComboStat1Cli.ItemIndex =4) Then Inc(Res3);
If (ComboStat2Cli.ItemIndex =4) Then Inc(Res3);
If (ComboStat3Cli.ItemIndex =4) Then Inc(Res3);
If (ComboStat4Cli.ItemIndex =4) Then Inc(Res3);

If ((app > 1) or (sect >1) or (res1 > 1) or (res2 > 1)or (res3 > 1) ) Then
    begin
    PGIBoxAf('Une valeur existe en double', TraduitGa('Stat client'));
    result:=FALSE;
    end;

End;

procedure TFParamCli.ComboStat2CliClick(Sender: TObject);
begin
TestStatUtilise;
if (TestStat = False) then TWinControl(ComboStat2Cli).SetFocus;
end;


procedure TFParamCli.ComboStat3CliClick(Sender: TObject);
begin
TestStatUtilise;
if (TestStat = False) then TWinControl(ComboStat3Cli).SetFocus;
end;

procedure TFParamCli.ComboStat4cliClick(Sender: TObject);
begin
TestStatUtilise;
if (TestStat = False) then TWinControl(ComboStat4Cli).SetFocus;
end;

procedure TFParamCli.FormShow(Sender: TObject);
begin
TestStatUtilise;
end;


function TFParamCli.TestTableLibre : Boolean;
var
gt: array[1..27]  of Integer;
nb : integer;
begin
          // cette fct permet de contrôler qu'une même grille technique
          // n'a pas été choisi 2 fois  dans 2 tables libres différentes
For Nb:= 1 to 27 do
     Begin
     gt[Nb]:=0;
     End;
if ComboTableLibre1.ItemIndex > 0 then inc(gt[ComboTableLibre1.ItemIndex]);
if ComboTableLibre2.ItemIndex > 0 then inc(gt[ComboTableLibre2.ItemIndex]);
if ComboTableLibre3.ItemIndex > 0 then inc(gt[ComboTableLibre3.ItemIndex]);
if ComboTableLibre4.ItemIndex > 0 then inc(gt[ComboTableLibre4.ItemIndex]);
if ComboTableLibre5.ItemIndex > 0 then inc(gt[ComboTableLibre5.ItemIndex]);
if ComboTableLibre6.ItemIndex > 0 then inc(gt[ComboTableLibre6.ItemIndex]);
if ComboTableLibre7.ItemIndex > 0 then inc(gt[ComboTableLibre7.ItemIndex]);
if ComboTableLibre8.ItemIndex > 0 then inc(gt[ComboTableLibre8.ItemIndex]);
if ComboTableLibre9.ItemIndex > 0 then inc(gt[ComboTableLibre9.ItemIndex]);
if ComboTableLibre10.ItemIndex > 0 then inc(gt[ComboTableLibre10.ItemIndex]);
result:=TRUE;
For Nb:= 1 to 27 do
     Begin
     If (gt[nb] > 1) Then
        begin
        PGIBoxAf('Une valeur existe en double', TraduitGa('Table libre'));
        result:=FALSE;
        end;
     End;


end;


procedure TFParamCli.ComboTableLibre1Click(Sender: TObject);
begin
if (TestTableLibre = False) then TWinControl(COmboTableLibre1).SetFocus
end;

procedure TFParamCli.ComboTableLibre2Click(Sender: TObject);
begin
if (TestTableLibre = False) then TWinControl(COmboTableLibre2).SetFocus

end;

procedure TFParamCli.ComboTableLibre3Click(Sender: TObject);
begin
if (TestTableLibre = False) then TWinControl(COmboTableLibre3).SetFocus

end;

procedure TFParamCli.ComboTableLibre4Click(Sender: TObject);
begin
if (TestTableLibre = False) then TWinControl(COmboTableLibre4).SetFocus

end;

procedure TFParamCli.ComboTableLibre5Click(Sender: TObject);
begin
if (TestTableLibre = False) then TWinControl(COmboTableLibre5).SetFocus

end;

procedure TFParamCli.ComboTableLibre6Click(Sender: TObject);
begin
if (TestTableLibre = False) then TWinControl(COmboTableLibre6).SetFocus

end;

procedure TFParamCli.ComboTableLibre7Click(Sender: TObject);
begin
if (TestTableLibre = False) then TWinControl(COmboTableLibre7).SetFocus

end;

procedure TFParamCli.ComboTableLibre8Click(Sender: TObject);
begin
if (TestTableLibre = False) then TWinControl(COmboTableLibre8).SetFocus

end;

procedure TFParamCli.ComboTableLibre9Click(Sender: TObject);
begin
if (TestTableLibre = False) then TWinControl(COmboTableLibre9).SetFocus

end;

procedure TFParamCli.ComboTableLibre10Click(Sender: TObject);
begin
if (TestTableLibre = False) then TWinControl(COmboTableLibre10).SetFocus

end;

end.
