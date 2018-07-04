unit AssistImportBtp;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, UTob,
  assist, ComCtrls, HSysMenu, hmsgbox, HTB97, StdCtrls, Hctrls, ExtCtrls, IniFiles,
  Grids, HPanel, UiUtil, MenuOlg, ImgList, RecupBtp, DBTables,
  Hqry, Db, Mask, Spin, Outline, DirOutln, filectrl, Menus;

type
  TFAssistimportBtp = class(TFAssist)
    TabSheet3: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet6: TTabSheet;
    TabSheet4: TTabSheet;
    TabSheet5: TTabSheet;
    StringGrid1: THgrid;
    OpenDialog1: TOpenDialog;
    Label1: TLabel;
    ImageList1: TImageList;
    Button1: TButton;
    Memo1: TMemo;
    Titre2_1: TLabel;
    BAssociationChamps: TButton;
    Bc_InitBase: TCheckBox;
    GB_inibase: TGroupBox;
    RB_Creation: TRadioButton;
    RB_Maj: TRadioButton;
    CB_inibase: TCheckBox;
    GB_compteclient: TGroupBox;
    RB_clientC: TRadioButton;
    RB_client9: TRadioButton;
    GB_comptefour: TGroupBox;
    RB_FourniF: TRadioButton;
    RB_Fourni0: TRadioButton;
    TCollectifCli: THLabel;
    TCB_inibase: THLabel;
    TRB_Creation: THLabel;
    TRB_Maj: THLabel;
    TDebutCompteCli: THLabel;
    TDebutCompteFou: THLabel;
    TCollectifFou: THLabel;
    Collectifcli: THCritMaskEdit;
    CollectifFou: THCritMaskEdit;
    PanelFin: TPanel;
    TTextFin1: THLabel;
    TTextFin2: THLabel;
    ListRecap: TListBox;
    TRecap: THLabel;
    RB_MajP: TRadioButton;
    HLabel8: THLabel;
    OpenDialog2: TOpenDialog;
    TabSheet7: TTabSheet;
    Panelbioenvenue: TPanel;
    HLabel3: THLabel;
    HLabel4: THLabel;
    GB_Repertoire: TGroupBox;
    TREP: THLabel;
    CB_Repertoire: THValComboBox;
    B_Repertoire: TButton;
    HLabel9: THLabel;
    CMContexte: THCritMaskEdit;
    GB_CodeBarre: TGroupBox;
    TQualifCodeBarre: THLabel;
    TJournalVente: THLabel;
    QualifCodeBarre: THValComboBox;
    CMJournalVente: THCritMaskEdit;
    GB_Monnaie: TGroupBox;
    TMonnaieTenue: THLabel;
    TPlafond: THLabel;
    TCommission: THLabel;
    RBMonnaieFRF: TRadioButton;
    RBMonnaieEURO: TRadioButton;
    NEPlafond: THNumEdit;
    CBCommission: THValComboBox;
    GB_Statistiques: TGroupBox;
    HLabel6: THLabel;
    HLabel7: THLabel;
    SE_Client: TSpinEdit;
    SE_Article: TSpinEdit;
    GB_Stocks: TGroupBox;
    HLabel2: THLabel;
    HLabel5: THLabel;
    DateDebutExercice: THCritMaskEdit;
    NbMois: THNumEdit;
    ComboOption: TComboBox;
    procedure FormShow(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    procedure StringGrid1Click(Sender: TObject);
    procedure FormClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure StringGrid1DblClick(Sender: TObject);
    procedure StringGrid1SelectCell(Sender: TObject; ACol, ARow: Integer; var CanSelect: Boolean);
    procedure BAssociationChampsClick(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure B_RepertoireClick(Sender: TObject);
    procedure StringGrid1MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure ComboOptionChange(Sender: TObject);
    procedure StringGrid1TopLeftChanged(Sender: TObject);
  private
    procedure RecupDonnee(st_chmin: string);
    procedure remplitGrid;
    procedure RemplitRecap;
    { Déclarations privées }
  public
    { Déclarations publiques }
    ModeParam : boolean;
  end;

  procedure Assist_ImportBtp (AvecParam : boolean);

var
    Tob_Detables, Tob_Dechamp : Tob;
    i_NumEcran : integer;
    st_Chemin : string;

implementation

uses AssocImportBtp;

{$R *.DFM}

Procedure Assist_ImportBtp (AvecParam : boolean) ;
var
   Fo_Assist : TFAssistImportBtp;
Begin
     Fo_Assist := TFAssistImportBtp.Create (Application);
     fo_Assist.ModeParam := AvecParam;
     Try
         Fo_Assist.ShowModal;
     Finally
         Fo_Assist.free;
     End;
end;

procedure TFAssistImportBtp.FormShow(Sender: TObject);
var
    i_ind1 : integer;
    Onglet : TTabSheet;
    St_NomOnglet : String;
    Ficliste : TIniFile;

begin
    inherited;
    bAnnuler.Visible := True;
    bFin.Visible := True;
    bFin.Enabled := False;
    //i_NumEcran := 0;

    Onglet := P.ActivePage;
    st_NomOnglet := Onglet.Name;
    i_NumEcran := strtoint (Copy (st_NomOnglet, length (st_NomOnglet), 1)) - 1;
    Image.ImageIndex := i_NumEcran;
    if not ModeParam  then
       begin
       //P.Pages[1].TabVisible := false;
       P.Pages[3].TabVisible := false;
       P.Pages[4].TabVisible := false;
       P.Pages[0].PageIndex := 0;
       P.Pages[1].PageIndex := 1;
       P.Pages[2].PageIndex := 2;
       P.Pages[5].PageIndex := 3;
       end;

    StringGrid1.ColWidths[0] := 100;
    StringGrid1.ColWidths[1] := 100;
    StringGrid1.ColWidths[2] := 100;
    StringGrid1.ColWidths[3] := 100;
    StringGrid1.ColWidths[4] := 100;
    StringGrid1.ColWidths[5] := 100;
    StringGrid1.ColWidths[StringGrid1.ColCount-1] := 100;
    if not (ModeParam) then st_chemin := CB_Repertoire.text
                       else st_Chemin := ExtractFilePath(ParamStr(0));

    if FileExists(st_Chemin + 'ListeRepertoire') then
    begin
        Ficliste := TIniFile.Create(st_Chemin + 'ListeRepertoire');
        Ficliste.ReadSectionValues('General', CB_Repertoire.Items.AnsiStrings);
        FicListe.Free;
        for i_ind1 := 0 to CB_Repertoire.Items.Count -1 do
            begin
            st_chemin := CB_Repertoire.Items.Strings[i_ind1];
            st_chemin := Copy(st_chemin, Pos('=', st_chemin) + 1, 255);
            CB_Repertoire.Items.Strings[i_ind1] := st_chemin;
            end;
    end
    else
    begin
    CB_Repertoire.Values.Add(st_Chemin);
    CB_Repertoire.Items.Add(st_Chemin);
    end;
    CB_Repertoire.ItemIndex := 0;
end;
//

procedure TFAssistImportBtp.RecupDonnee (st_chmin : string);
var
    i_ind1, i_ind2,i_indliste : integer;
    Onglet : TTabSheet;
    St_NomOnglet : String;
    st_ident, st_valeur : string;
    st_chaine : String;
    Ficliste : TIniFile;
    st_enreg : String;
    Ts_listeFic : Tstrings;
begin
Ficliste := TIniFile.Create(st_Chemin + 'ListeFic');
st_enreg := Ficliste.ReadString ('General', 'DelFic', '');
if st_enreg = 'O' then CB_inibase.checked := True
                  else CB_inibase.checked := False;
st_enreg := Ficliste.ReadString ('General', 'ModeCreation', '');
if st_enreg = 'O' then
begin
    RB_Creation.checked := True;
end
else if st_enreg = 'M' then
begin
    RB_Maj.checked := True;
end
else
begin
    RB_MajP.checked := True;
end;
st_enreg := Ficliste.ReadString ('General', 'Client', '');
if st_enreg = 'C' then
begin
    RB_clientC.checked := True;
end
else
begin
    RB_client9.checked := True;
end;
st_enreg := Ficliste.ReadString ('General', 'Fournisseur', '');
if st_enreg = 'F' then
begin
    RB_FourniF.checked := True;
end
else
begin
    RB_Fourni0.checked := True;
end;
st_enreg := Ficliste.ReadString ('General', 'Monnaie', '');
if st_enreg = 'F' then
begin
    RBMonnaieFRF.checked := True;
end
else
begin
    RBMonnaieEURO.checked := True;
end;
CBCommission.Text := Ficliste.ReadString ('General', 'CommissionText', '');
CBCommission.Value := Ficliste.ReadString ('General', 'CommissionValue', '');
NEPlafond.Value := Ficliste.ReadFloat ('General', 'Plafond', 0.0);
QualifCodeBarre.Text := Ficliste.ReadString ('General', 'QualifBarreText', '');
QualifCodeBarre.Value := Ficliste.ReadString ('General', 'QualifBarreVal', '');
CMJournalVente.Text := Ficliste.ReadString ('General', 'JournalVente', '');
CollectifCli.Text := Ficliste.ReadString ('General', 'ColClient', '');
CollectifFou.Text := Ficliste.ReadString ('General', 'ColFourni', '');
SE_Client.Value := Ficliste.ReadInteger ('General', 'NoStatClient', 0);
SE_Article.Value := Ficliste.ReadInteger ('General', 'NoStatArticle', 0);

Ts_listeFic := TstringList.Create;
i_ind1 := 0;
Ts_listeFic.Text := ''; // Réinitialisation de la TString
Ficliste.ReadSectionValues ('Fichier', Ts_listeFic);
RemplitGrid;
if Ts_ListeFic.count <> 0 then
   begin
   while i_ind1 <= Ts_listeFic.Count - 1 do
      begin
      st_enreg := Ts_listeFic.Strings [i_ind1];
      ST_CHAINE := Copy(st_enreg, 0, Pos('=', st_enreg) - 1);
      st_enreg := Copy(st_enreg, Pos('=', st_enreg) + 1, 255);
      for i_indliste := 0 to StringGrid1.RowCount -1 do
          begin
          if StringGrid1.cells[0,i_indliste] = st_chaine then break;
          end;
      i_ind2 := 1;
      if i_ind2 <= StringGrid1.ColCount then
         begin
         while Pos(';', st_enreg) <> 0 do
             begin
             if Copy(st_enreg, 0, Pos(';', st_enreg) - 1) <> '' then
                StringGrid1.Cells[i_ind2, i_indListe] := Copy(st_enreg, 0, Pos(';', st_enreg) - 1);
             st_enreg := Copy(st_enreg, Pos(';', st_enreg) + 1, 255);
             i_ind2 := i_ind2 + 1;
             end;
         // StringGrid1.Cells[i_ind2, i_ind1] := st_enreg;
         end;
      i_ind1 := i_ind1 + 1;
      end;
   end;
Ts_listeFic.Free;
Ficliste.Free;
end;

procedure TFAssistImportBtp.remplitGrid;
var TOB_DeTables : TOB;
    Indice : integer;
begin
Tob_Detables := TOB.Create('DETABLES', nil, -1);
Tob_Detables.LoadDetailDB('DETABLES','','',Nil,FALSE);
StringGrid1.RowCount := Tob_Detables.Detail.Count;
Tob_Detables.PutStringsDetail(StringGrid1.Cols[0],'DT_NOMTABLE');
Tob_Detables.Free;
for Indice := 0 to stringGrid1.rowcount -1 do
    begin
    StringGrid1.Cells[StringGrid1.ColCount -1, Indice] := 'Défaut';
    end;
end;

//
procedure TFAssistImportBtp.RemplitRecap;
var st_chaine : String;
begin
ListRecap.Items.Clear;
ListRecap.Items.Add ('---> ' + GB_inibase.Caption);
if CB_inibase.checked then
begin
    st_chaine := ' : Oui';
end
else
begin
    st_chaine := ' : Non';
end;
ListRecap.Items.Add (CB_inibase.Caption +
                     st_chaine);
if RB_Creation.Checked Then
begin
    ListRecap.Items.Add (RB_Creation.Caption);
end
else
begin
    ListRecap.Items.Add (RB_Maj.Caption);
end;
ListRecap.Items.Add ('');
ListRecap.Items.Add ('---> ' + GB_CodeBarre.Caption);
ListRecap.Items.Add (TQualifCodeBarre.Caption + ' : ' +
                     QualifCodeBarre.Text);
ListRecap.Items.Add ('');
ListRecap.Items.Add ('---> ' + GB_compteclient.Caption);
if RB_clientC.Checked then
begin
    st_chaine := ' : "C"';
end
else
begin
    st_chaine := ' : "9"';
end;
ListRecap.Items.Add (TDebutCompteCli.Caption +
                     st_chaine);
ListRecap.Items.Add (TCollectifCli.Caption + ' : ' +
                     CollectifCli.Text);
ListRecap.Items.Add (TJournalVente.Caption + ' : ' +
                     CMJournalVente.Text);
ListRecap.Items.Add ('');
ListRecap.Items.Add ('---> ' + GB_comptefour.Caption);
if RB_FourniF.Checked then
begin
    st_chaine := ' : "F"';
end
else
begin
    st_chaine := ' : "0"';
end;
ListRecap.Items.Add (TDebutCompteFou.Caption +
                     st_chaine);
ListRecap.Items.Add (TCollectifFou.Caption + ' : ' +
                     CollectifFou.Text);
ListRecap.Items.Add ('');
ListRecap.Items.Add ('---> ' + GB_Monnaie.Caption);
if RBMonnaieFRF.Checked then
begin
    st_chaine := ' : Francs';
end
else
begin
    st_chaine := ' : Euros';
end;
ListRecap.Items.Add (TMonnaieTenue.Caption + st_chaine);
ListRecap.Items.Add (TPlafond.Caption + ' : ' + FloatToStr(NEPlafond.Value));
ListRecap.Items.Add (TCommission.Caption + ' : ' + CBCommission.Text);
ListRecap.Items.Add ('');
end;
//
procedure TFAssistImportBtp.bSuivantClick(Sender: TObject);
var
    i_ind1, i_ind2 : integer;
    Onglet : TTabSheet;
    St_NomOnglet : String;
    st_ident, st_valeur : string;
    st_chaine : String;
    Ficliste : TIniFile;
    st_enreg : String;
    Ts_listeFic : Tstrings;
begin
    Onglet := P.ActivePage;
    st_NomOnglet := Onglet.Name;
    i_NumEcran := strtoint (Copy (st_NomOnglet, length (st_NomOnglet), 1)) - 1;
    inherited;
    case i_NumEcran of
    4 :
        begin
            // bFin.Enabled := True;
            Ficliste := TIniFile.Create(st_Chemin + 'ListeFic');
            Ficliste.EraseSection ('General');
            if CB_inibase.checked then
            begin
                Ficliste.WriteString ('General', 'DelFic', 'O');
            end
            else
            begin
                Ficliste.WriteString ('General', 'DelFic', 'N');
            end;
            if RB_Creation.checked then
            begin
                Ficliste.WriteString ('General', 'ModeCreation', 'O');
            end
            else if RB_Maj.checked then
            begin
                Ficliste.WriteString ('General', 'ModeCreation', 'M');
            end
            else
            begin
                Ficliste.WriteString ('General', 'ModeCreation', 'P');
            end;
            if RB_clientC.checked then
            begin
                Ficliste.WriteString ('General', 'Client', '"C"');
            end
            else
            begin
                Ficliste.WriteString ('General', 'Client', '"9"');
            end;
            if RB_FourniF.checked then
            begin
                Ficliste.WriteString ('General', 'Fournisseur', '"F"');
            end
            else
            begin
                Ficliste.WriteString ('General', 'Fournisseur', '"0"');
            end;
            if RBMonnaieFRF.checked then
            begin
                Ficliste.WriteString ('General', 'Monnaie', 'F');
            end
            else
            begin
                Ficliste.WriteString ('General', 'Monnaie', 'E');
            end;
            Ficliste.WriteString ('General', 'Plafond', FloatToStr(NEPlafond.Value));
            Ficliste.WriteString ('General', 'CommissionText', CBCommission.Text);
            Ficliste.WriteString ('General', 'CommissionValue', CBCommission.Value);
            Ficliste.WriteString ('General', 'NoStatClient', IntToStr(SE_Client.Value));
            Ficliste.WriteString ('General', 'QualifBarreText', QualifCodeBarre.Text);
            Ficliste.WriteString ('General', 'QualifBarreVal', QualifCodeBarre.Value);
            Ficliste.WriteString ('General', 'JournalVente', CMJournalVente.Text);
            Ficliste.WriteString ('General', 'ColClient', CollectifCli.Text);
            Ficliste.WriteString ('General', 'ColFourni', CollectifFou.Text);
            Ficliste.WriteString ('General', 'NoStatClient', IntToStr(SE_Client.Value));
            Ficliste.WriteString ('General', 'NoStatArticle', IntToStr(SE_Article.Value));
            Ficliste.WriteString ('General', 'DebutExercice', DateDebutExercice.Text);
            Ficliste.WriteString ('General', 'NbMois', NbMois.Text);
            FicListe.Free;
        end;
    1 :
        begin
        RemplitRecap;
        end;
    2 :
        begin
        Ficliste := TIniFile.Create(st_Chemin + 'ListeFic');
        Ficliste.EraseSection ('Fichier');
        i_ind1 := 0;
        while i_ind1 <= StringGrid1.RowCount do
            begin
            if (StringGrid1.Cells[0, i_ind1] <> '') and (stringgrid1.cells[1,i_ind1] <> '') then
               begin
               st_ident := StringGrid1.Cells[0, i_ind1];
               st_valeur := '';
               for i_ind2 := 1 to stringGrid1.colcount -2 do
                   begin
                   st_valeur := st_valeur + StringGrid1.Cells[i_ind2, i_ind1] + ';'
                   end;
               if copy(st_valeur,1,1) <> ';' then
                  begin
                  if StringGrid1.Cells[StringGrid1.colCount -1,i_ind1 ] <> 'Défaut' then
                     st_valeur := st_valeur + StringGrid1.Cells[StringGrid1.colCount -1, i_ind1]+';';
                  FicListe.WriteString('Fichier', st_ident, st_valeur);
                  end;
               end;
            i_ind1 := i_ind1 + 1;
            end;
        FicListe.Free;
        end;
    6 :
        begin
            Ficliste := TIniFile.Create(ExtractFilePath(ParamStr(0)) + 'ListeRepertoire');
            for i_ind1 := 0 to CB_Repertoire.Items.Count - 1 do
                begin
                st_enreg := CB_Repertoire.Items.Strings[i_ind1];
                Ficliste.WriteString ('General', 'Repertoire'+IntToStr(i_ind1), st_enreg);
                end;
            FicListe.Free;
            st_Chemin := CB_Repertoire.Text;
            if FileExists(st_Chemin + 'ListeFic') then
            begin
                RecupDonnee (st_chemin);
            end
            else
            begin
               remplitGrid;
            end;
        end;
    end;
    Onglet := P.ActivePage;
    st_NomOnglet := Onglet.Name;
    i_NumEcran := strtoint (Copy (st_NomOnglet, length (st_NomOnglet), 1)) - 1;
    //i_NumEcran := i_NumEcran + 1;
    Image.ImageIndex := i_NumEcran;
    if (P.ActivePageIndex = 3) and (not ModeParam ) then
       begin
       Bsuivant.Enabled := false;
       RemplitRecap;
       end;
    if (bSuivant.Enabled) then
    begin
        bFin.Enabled := False;
    end
    else
    begin
        bFin.Enabled := True;
    end;
end;

procedure TFAssistImportBtp.Button1Click(Sender: TObject);
var chaine : string;
    i : integer;
begin
    inherited;
    OpenDialog1.InitialDir := CB_Repertoire.Text;
    OpenDialog1.Filter := 'Descriptifs PGI (*.des)|*.des';
    if OpenDialog1.Execute then
       if OpenDialog1.FileName <> '' then
          begin
          StringGrid1.Cells[StringGrid1.Col, StringGrid1.Row] := Copy(OpenDialog1.filename, 0, Pos('.', OpenDialog1.filename) - 1);
          end;
//    else
  //      StringGrid1.Cells[StringGrid1.Col, StringGrid1.Row] := '';
    Button1.Visible := False;
end;

procedure TFAssistImportBtp.bPrecedentClick(Sender: TObject);
Var
    Onglet : TTabSheet;
    St_NomOnglet : String;
begin
    inherited;
    Onglet := P.ActivePage;
    st_NomOnglet := Onglet.Name;
    i_NumEcran := strtoint (Copy (st_NomOnglet, length (st_NomOnglet), 1)) - 1;
    //i_NumEcran := i_NumEcran - 1;
    Image.ImageIndex := i_NumEcran;
    if (bSuivant.Enabled) then
    begin
        bFin.Enabled := False;
    end
    else
    begin
        bFin.Enabled := True;
    end;
    //if i_NumEcran = 0 then bFin.Enabled := False;
end;

procedure TFAssistImportBtp.StringGrid1Click(Sender: TObject);
var
    Sel_Rect : TRect;
    index : integer;
begin
  inherited;
//    Button1.Visible := False;
    if StringGrid1.col < StringGrid1.colcount -1 then
       begin
       Sel_Rect := StringGrid1.CellRect(StringGrid1.Col, StringGrid1.Row);
       ComboOption.Visible := false;
       Button1.Top := Sel_Rect.Top + StringGrid1.Top + 2;
       Button1.Left := Sel_Rect.Right - Button1.Width + StringGrid1.Left + 2;
       Button1.Height := Sel_Rect.Bottom - Sel_Rect.Top;
       Button1.Visible := True;
       end else
       begin
       Sel_Rect := StringGrid1.CellRect(StringGrid1.Col, StringGrid1.Row);
       Button1.Visible := false;
       for Index := 0 to ComboOption.Items.Count -1 do
           if ComboOption.items.Strings[Index]= StringGrid1.Cells[StringGrid1.Col, StringGrid1.Row] then
              begin
              ComboOption.ItemIndex := Index;
              break;
              end;
       ComboOption.Top := Sel_Rect.Top+ StringGrid1.Top + 2;
       ComboOption.Left := Sel_Rect.Left+ StringGrid1.Left + 2;
       ComboOption.Visible := true;
       end;
end;

procedure TFAssistImportBtp.FormClick(Sender: TObject);
begin
  inherited;
    Button1.Visible := False;
    ComboOption.Visible := False;
end;

procedure TFAssistImportBtp.FormCreate(Sender: TObject);
begin
  inherited;
    FAssocImportBTP := TFAssocImportBTP.Create(FAssocImportBTP);
end;

procedure TFAssistImportBtp.StringGrid1DblClick(Sender: TObject);
begin
  inherited;
  if StringGrid1.Col < StringGrid1.colCount then Button1Click(Sender);
end;

procedure TFAssistImportBtp.StringGrid1SelectCell(Sender: TObject; ACol,
  ARow: Integer; var CanSelect: Boolean);
var
    Sel_Rect : TRect;
    Index : integer;
begin
  inherited;
  if Acol < StringGrid1.colcount -1 then
     begin
     Sel_Rect := StringGrid1.CellRect(StringGrid1.Col, StringGrid1.Row);
    ComboOption.Visible := false;
    Button1.Top := Sel_Rect.Top + StringGrid1.Top + 2;
    Button1.Left := Sel_Rect.Right - Button1.Width + StringGrid1.Left + 2;
    Button1.Height := Sel_Rect.Bottom - Sel_Rect.Top;
    Button1.Visible := True;
    end else
    begin
    Sel_Rect := StringGrid1.CellRect(ACol, ARow);
    Button1.Visible := false;
    for Index := 0 to ComboOption.Items.Count -1 do
        if ComboOption.items.Strings[Index]= StringGrid1.Cells [ACol, ARow] then
           begin
           ComboOption.ItemIndex := Index;
           break;
           end;
    ComboOption.Top := Sel_Rect.Top + StringGrid1.top + 2;
    ComboOption.Left := Sel_Rect.left+ StringGrid1.left +2;
    ComboOption.Visible := true;
    end;

end;

procedure TFAssistImportBtp.BAssociationChampsClick(Sender: TObject);
var
    i_ind1, i_ind2 : integer;
begin
    inherited;

    i_ind2 := 0;
    for i_ind1 := 0 to stringgrid1.rowcount   do
        begin
            if Stringgrid1.cells[1,i_ind1] <> '' then
               begin
               fassocimportBTP.stringgrid1.cells[0,i_ind2] := stringgrid1.cells[0,i_ind1];
               fassocimportBTP.stringgrid1.cells[1,i_ind2] := stringgrid1.cells[1,i_ind1];
               fassocimportBTP.stringgrid1.cells[2,i_ind2] := stringgrid1.cells[2,i_ind1];
               fassocimportBTP.stringgrid1.cells[3,i_ind2] := stringgrid1.cells[3,i_ind1];
               fassocimportBTP.stringgrid1.cells[4,i_ind2] := stringgrid1.cells[4,i_ind1];
//               fassocimport.stringgrid1.cells[5,i_ind2] := stringgrid1.cells[5,i_ind1];
               i_ind2 := i_ind2 + 1;
               end;
        end;
    fassocimportBTP.stringgrid1.rowcount := i_ind2;
    fassocimportBTP.stringgrid1.width := stringgrid1.width;
    fassocimportBTP.stringgrid1.colwidths[0] := stringgrid1.colwidths[0];
    fassocimportBTP.stringgrid1.colwidths[1] := stringgrid1.colwidths[1];
    fassocimportBTP.stringgrid1.colwidths[2] := stringgrid1.colwidths[2];
    fassocimportBTP.stringgrid1.colwidths[3] := stringgrid1.colwidths[3];
    fassocimportBTP.stringgrid1.colwidths[4] := stringgrid1.colwidths[4];
    fassocimportBTP.stringgrid1.colwidths[5] := stringgrid1.colwidths[5];
    //fassocimport.stringgrid1.colwidths[6] := stringgrid1.colwidths[6];
    fassocimportBTP.WorkingPath.Text := st_Chemin;
    fassocimportBTP.showmodal;
end;

procedure TFAssistImportBtp.bFinClick(Sender: TObject);
begin
    inherited;
    RecuperationBtp (st_Chemin);
    Close ;
end;


procedure TFAssistImportBtp.B_RepertoireClick(Sender: TObject);
var
   Directory : string;
   HelpCtx: Longint;

begin
  inherited;
  if SelectDirectory(Directory, [sdAllowCreate, sdPrompt], HelpCtx) then
     Begin
     CB_Repertoire.Value := Directory + '\';
     CB_Repertoire.Values.Add(Directory + '\');
     CB_Repertoire.Items.Add(Directory + '\');
     CB_Repertoire.Text := Directory + '\';
     SetCurrentDir(Directory);
     end;
end;

procedure TFAssistImportBtp.StringGrid1MouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
var Acol,Arow : integer;
    SResult : string;
begin
  inherited;
  StringGrid1.MouseToCell (X,Y,Acol,Arow);
  if (Acol >= 0 )and (Acol <= StringGrid1.colcount) and (StringGrid1.Cells [Acol,Arow] <> '') then
     begin
     if Acol = 0 then Sresult := 'Table '
                 else Sresult := 'Fichier ';
     Sresult := Sresult + StringGrid1.Cells [Acol,Arow];
     stringGrid1.Hint := Sresult;
     StringGrid1.ShowHint := true;
     end else StringGrid1.ShowHint := false;

end;

procedure TFAssistImportBtp.ComboOptionChange(Sender: TObject);
begin
  inherited;
  StringGrid1.Cells[StringGrid1.Col, StringGrid1.Row] := ComboOption.Items.Strings [ComboOption.itemindex];
end;

procedure TFAssistImportBtp.StringGrid1TopLeftChanged(Sender: TObject);
begin
  inherited;
ComboOption.visible := false;
Button1.Visible := false;
end;

end.

