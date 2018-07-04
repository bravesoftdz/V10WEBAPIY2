unit SaisieConditionBTP;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Hctrls, Buttons, hmsgbox,
  AssistUtilBtp, HTB97, UTOB, ComCtrls;

Function  Entree_ConditionBtp(Parms: array of variant; nb: integer) : string;

type
  TFSaisieConditionBtp = class(TForm)
    HLabel1: THLabel;
    HLabel2: THLabel;
    HLabel3: THLabel;
    Condition: TMemo;
    Vrai: TMemo;
    Faux: TMemo;
    ListeTable: TListBox;
    ListeFichier: TListBox;
    ListeOper: THValComboBox;
    Label3: TLabel;
    Label2: TLabel;
    HLabel4: THLabel;
    bET: TButton;
    bOU: TButton;
    bOuvrePar: TButton;
    bFermePar: TButton;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    BitBtn3: TBitBtn;
    BitBtn4: TBitBtn;
    BitBtn5: TBitBtn;
    BitBtn6: TBitBtn;
    BitBtn7: TBitBtn;
    BitBtn8: TBitBtn;
    BitBtn9: TBitBtn;
    BitBtn10: TBitBtn;
    BitBtn11: TBitBtn;
    BitBtn12: TBitBtn;
    BitBtn13: TBitBtn;
    BitBtn14: TBitBtn;
    BitBtn15: TBitBtn;
    MsgBox: THMsgBox;
    Dock971: TDock97;
    ToolWindow971: TToolWindow97;
    BAide: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BValider: TToolbarButton97;
    bValiderSi: TToolbarButton97;
    bValiderAlors: TToolbarButton97;
    bValiderSinon: TToolbarButton97;
    bSI: TButton;
    ListeTableCode: TListBox;
    ListeFichierCode: TListBox;
    TreeView2: TTreeView;
    bSuivant: TButton;
    bVrai: TButton;
    bFaux: TButton;
    BitBtn16: TBitBtn;
    HLabel5: THLabel;
    FctsDisp: THValComboBox;
    BVide: TButton;
    procedure FormShow(Sender: TObject);
    procedure ConditionClick(Sender: TObject);
    procedure VraiClick(Sender: TObject);
    procedure FauxClick(Sender: TObject);
    procedure ListeTableDblClick(Sender: TObject);
    procedure ListeFichierDblClick(Sender: TObject);
    procedure ListeOperChange(Sender: TObject);
    procedure ButtonClick(Sender: TObject);
    procedure bValiderSiClick(Sender: TObject);
    procedure bValiderAlorsClick(Sender: TObject);
    procedure bValiderSinonClick(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure bSIClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FctsDispChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Déclarations privées }
    TobCond, TSI, TALORS, TSINON : TOB;
    MA : TMemo;
    NomsTable, NomsFichier : string;
    procedure InsereValeur(Valeur : string);
  public
    { Déclarations publiques }
  end;

var
  FSaisieConditionBtp: TFSaisieConditionBtp;
  Retour : string;

implementation

{$R *.DFM}

function  Entree_ConditionBtp(Parms: array of variant; nb: integer) : string;
begin
FSaisieConditionBTP := TFSaisieConditionBtp.Create(Application);
FSaisieConditionBTP.ListeTable.Items.Text := String(Parms[0]);
FSaisieConditionBTP.ListeFichier.Items.Text := String(Parms[1]);
FSaisieConditionBTP.ListeTableCode.Items.Text := String(Parms[2]);
FSaisieConditionBTP.ListeFichierCode.Items.Text := String(Parms[3]);
FSaisieConditionBTP.Condition.Text := String(Parms[4]);
FSaisieConditionBTP.Vrai.Text := '';
FSaisieConditionBTP.Faux.Text := '';
FSaisieConditionBTP.ShowModal;
Result := Retour;
end;

procedure TFSaisieConditionBtp.FormShow(Sender: TObject);
var
    i_ind1, PosSi, PosAlors, PosSinon : integer;
    st1, st2, st3 : string;

begin
NomsTable := '';
for i_ind1 := 0 to ListeTableCode.Items.Count - 1 do
    NomsTable := NomsTable + ListeTableCode.Items.Strings[i_ind1] + ';';
NomsFichier := '';
for i_ind1 := 0 to ListeFichierCode.Items.Count - 1 do
    NomsFichier := NomsFichier + ListeFichierCode.Items.Strings[i_ind1] + ';';
TobCond := TOB.Create('', nil, -1);
TobCond.AddChampSup('Valeur', False);
TobCond.PutValue('Valeur', 'Détail');
if Pos('{SI', Condition.Text) <> 0 then
    begin
    Eclate_Si(Condition.Text, st1, st2, st3);
    Condition.Text := st1;
    Vrai.Text := st2;
    Faux.Text := st3;
    if Condition.Text <> '' then bValiderSiClick(Sender);
    if Vrai.Text <> ''      then bValiderAlorsClick(Sender);
    if Faux.Text <> ''      then bValiderSinonClick(Sender);
    TobCond.PutTreeView(TreeView2, nil, 'Valeur;Valeur;Valeur;Valeur;Valeur;Valeur;Valeur;Valeur;Valeur;Valeur;Valeur');
    end;
Condition.SetFocus;
MA := Condition;
bSI.Enabled := False;
TreeView2.FullExpand;
end;

procedure TFSaisieConditionBtp.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if ModalResult = mrCancel then
    retour := ''
    else
    begin
    Retour := '{SI ' + Condition.Text + ' ALORS ' + Vrai.Text;
    if Faux.Text <> '' then Retour := Retour + ' SINON ' + Faux.Text;
    Retour := Retour + '}';
    end;
end;

procedure TFSaisieConditionBtp.ConditionClick(Sender: TObject);
begin
MA := Condition;
bSI.Enabled := False;
bSuivant.Enabled := False;
FctsDisp.Enabled := False;
end;

procedure TFSaisieConditionBtp.VraiClick(Sender: TObject);
begin
MA := Vrai;
bSI.Enabled := True;
bSuivant.Enabled := True;
FctsDisp.Enabled := True;
end;

procedure TFSaisieConditionBtp.FauxClick(Sender: TObject);
begin
MA := Faux;
bSI.Enabled := True;
bSuivant.Enabled := True;
FctsDisp.Enabled := True;
end;

procedure TFSaisieConditionBtp.ListeTableDblClick(Sender: TObject);
begin
InsereValeur(ListeTableCode.Items.Strings[ListeTable.ItemIndex]);
end;

procedure TFSaisieConditionBtp.ListeFichierDblClick(Sender: TObject);
begin
InsereValeur(ListeFichierCode.Items.Strings[ListeFichier.ItemIndex]);
end;

procedure TFSaisieConditionBtp.ListeOperChange(Sender: TObject);
begin
InsereValeur(ListeOper.Value);
end;

procedure TFSaisieConditionBtp.FctsDispChange(Sender: TObject);
var
    st1 : string;

begin
st1 := FctsDisp.Value;
if st1 <> 'COMPTEUR' then
    begin
    st1 := st1 + '(';
    if st1 = 'REPLACECAR' then
        MA.Hint := 'REPLACECAR( Champ(s) à traiter, Chaine à remplacer, Chaine de remplacement )';
    if st1 = 'CORRESPOND' then
        MA.Hint := 'CORRESPOND( Champ(s) à traiter )';
    if st1 = 'MAJUSCULE' then
        MA.Hint := 'MAJUSCULE( Champ(s) à traiter )';
    if st1 = 'TABLE' then
        MA.Hint := 'TABLE( Table à lire )';
    if st1 = 'SOUSCHAINE' then
        MA.Hint := 'SOUSCHAINE( Champ à traiter, Début, Longueur )';
    if st1 = 'SPACE' then
        MA.Hint := 'SPACE( Nombre d''espaces à insérer )';
    if st1 = 'INTSTR' then
        MA.Hint := 'INTSTR( Champ à convertir )';
    if st1 = 'STRINT' then
        MA.Hint := 'STRINT( Champ à convertir )';
    end;
InsereValeur(st1);
end;

procedure TFSaisieConditionBtp.ButtonClick(Sender: TObject);
var
    Valeur : string;

begin
if Sender = bET       then Valeur := ' ET ';
if Sender = bOU       then Valeur := ' OU ';
if Sender = bOuvrePar then Valeur := '(';
if Sender = bFermePar then Valeur := ')';
if Sender = bSuivant  then Valeur := '"SUIVANT"';
if Sender = bVrai     then Valeur := '"VRAI"';
if Sender = bFaux     then Valeur := '"FAUX"';
if Sender = bVide     then Valeur := '"VIDE"';
if Sender = BitBtn1   then Valeur := '1';
if Sender = BitBtn2   then Valeur := '2';
if Sender = BitBtn3   then Valeur := '3';
if Sender = BitBtn4   then Valeur := '4';
if Sender = BitBtn5   then Valeur := '5';
if Sender = BitBtn6   then Valeur := '6';
if Sender = BitBtn7   then Valeur := '7';
if Sender = BitBtn8   then Valeur := '8';
if Sender = BitBtn9   then Valeur := '9';
if Sender = BitBtn10  then Valeur := '0';
if Sender = BitBtn11  then Valeur := '+';
if Sender = BitBtn12  then Valeur := '-';
if Sender = BitBtn13  then Valeur := '*';
if Sender = BitBtn14  then Valeur := '/';
if Sender = BitBtn15  then Valeur := '.';
if Sender = BitBtn16  then Valeur := '^';
InsereValeur(Valeur);
end;

procedure TFSaisieConditionBtp.InsereValeur(Valeur : string);
begin
MA.Text := Copy(MA.Text, 0, MA.SelStart) +
           Valeur +
           Copy(MA.Text, MA.SelStart + 1, Length(MA.Text) - 1);
MA.SelStart := Length(MA.Text);
MA.SetFocus;
end;

procedure TFSaisieConditionBtp.bValiderSiClick(Sender: TObject);
var
    Mess : string;

begin
if TSI = nil then
    begin
    TSI := TOB.Create('', TobCond, 0);
    TSI.AddChampSup('Type', True);
    TSI.AddChampSup('Valeur', True);
    TSI.PutValue('Type', 'Titre');
    TSI.PutValue('Valeur', 'SI');
    end
    else
    begin
    TSI.Detail.Clear;
    end;
if not Analyse_Condition(Condition.Text, NomsTable, NomsFichier, Mess, TSI) then
    begin
    MsgBox.Execute(0, Mess, '');
    Condition.SetFocus;
    ModalResult := 0;
    end;
TobCond.PutTreeView(TreeView2, nil, 'Valeur;Valeur;Valeur;Valeur;Valeur;Valeur;Valeur;Valeur;Valeur;Valeur;Valeur');
TreeView2.FullExpand;
end;

procedure TFSaisieConditionBtp.bValiderAlorsClick(Sender: TObject);
var
    Mess : string;

begin
if TALORS = nil then
    begin
    TALORS := TOB.Create('', TobCond, 1);
    TALORS.AddChampSup('Type', True);
    TALORS.AddChampSup('Valeur', True);
    TALORS.PutValue('Type', 'Titre');
    TALORS.PutValue('Valeur', 'ALORS');
    end
    else
    begin
    TALORS.Detail.Clear;
    end;
if not Analyse_Vrai_Faux(Vrai.Text, NomsTable, NomsFichier, Mess, TALORS) then
    begin
    MsgBox.Execute(0, Caption, '');
    Vrai.SetFocus;
    ModalResult := 0;
    end;
TobCond.PutTreeView(TreeView2, nil, 'Valeur;Valeur;Valeur;Valeur;Valeur;Valeur;Valeur;Valeur;Valeur;Valeur;Valeur');
TreeView2.FullExpand;
end;

procedure TFSaisieConditionBtp.bValiderSinonClick(Sender: TObject);
var
    Mess : string;

begin
if TSINON = nil then
    begin
    TSINON := TOB.Create('', TobCond, 2);
    TSINON.AddChampSup('Type', True);
    TSINON.AddChampSup('Valeur', True);
    TSINON.PutValue('Type', 'Titre');
    TSINON.PutValue('Valeur', 'SINON');
    end
    else
    begin
    TSINON.Detail.Clear;
    end;
if not Analyse_Vrai_Faux(Faux.Text, NomsTable, NomsFichier, Mess, TSINON) then
    begin
    MsgBox.Execute(0, Caption, '');
    Faux.SetFocus;
    ModalResult := 0;
    end;
TobCond.PutTreeView(TreeView2, nil, 'Valeur;Valeur;Valeur;Valeur;Valeur;Valeur;Valeur;Valeur;Valeur;Valeur;Valeur');
TreeView2.FullExpand;
end;

procedure TFSaisieConditionBtp.BValiderClick(Sender: TObject);
begin
bValiderSiClick(Sender);
bValiderAlorsClick(Sender);
bValiderSinonClick(Sender);
TobCond.PutTreeView(TreeView2, nil, 'Valeur;Valeur;Valeur;Valeur;Valeur;Valeur;Valeur;Valeur;Valeur;Valeur;Valeur');
TreeView2.FullExpand;
end;

procedure TFSaisieConditionBtp.BAbandonClick(Sender: TObject);
begin
Condition.Text := '';
Vrai.Text := '';
Faux.Text := '';
end;

procedure TFSaisieConditionBtp.bSIClick(Sender: TObject);
var
    i_ind1, i_ind2, i_ind3 : integer;
    st1, st2, SiAlorsSinon, sCondition, sVrai, sFaux : string;
    Table, Fichier, TableCode, FichierCode : string;
begin
if (MA <> Vrai) and (MA <> Faux) then Exit;
Table := ListeTable.Items.Text;
Fichier := ListeFichier.Items.Text;
TableCode := ListeTableCode.Items.Text;
FichierCode := ListeFichierCode.Items.Text;
SiAlorsSinon := '';
if Pos('{', MA.Text) <> 0 then
    begin
    for i_ind1 := MA.SelStart downto 0 do
        if MA.Text[i_ind1] = '{' then Break;
    if i_ind1 >= 1 then
        begin
        MA.SelStart := i_ind1 - 1;
        i_ind2 := 0;
        for i_ind3 := i_ind1 to Length(MA.Text) do
            begin
            SiAlorsSinon := SiAlorsSinon + MA.Text[i_ind3];
            if MA.Text[i_ind3] = '{' then Inc(i_ind2);
            if MA.Text[i_ind3] = '}' then Dec(i_ind2);
            if i_ind2 = 0 then Break;
            end;
        MA.SelLength := i_ind3 - i_ind1 + 1;
        end;
    end;
st1 := Entree_ConditionBtp([Table, Fichier, TableCode, FichierCode, SiAlorsSinon], 5);
MA.SelText := st1;
//FSaisieCondition.Free;
end;

end.
