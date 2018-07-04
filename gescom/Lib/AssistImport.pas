unit AssistImport;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, UTob,
  assist, ComCtrls, HSysMenu, hmsgbox, HTB97, StdCtrls, Hctrls, ExtCtrls, IniFiles,
  Grids, HPanel, UiUtil, MenuOlg, AssocImport, ImgList, RecupNeg,
{$IFNDEF EAGLCLIENT}
     DB,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  Hqry, Mask, Spin, Outline, DirOutln, filectrl, EclateTra, HEnt1,
  Buttons;

type
  TFAssistimport = class(TFAssist)
    TabSheet3: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet6: TTabSheet;
    TabSheet4: TTabSheet;
    TabSheet5: TTabSheet;
    StringGrid1: TStringGrid;
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
    TabSheet1: TTabSheet;
    GB_CodeBarre: TGroupBox;
    TQualifCodeBarre: THLabel;
    QualifCodeBarre: THValComboBox;
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
    GB_Statistiques: TGroupBox;
    HLabel6: THLabel;
    HLabel7: THLabel;
    SE_Client: TSpinEdit;
    SE_Article: TSpinEdit;
    GB_Monnaie: TGroupBox;
    TMonnaieTenue: THLabel;
    TPlafond: THLabel;
    TCommission: THLabel;
    RBMonnaieFRF: TRadioButton;
    RBMonnaieEURO: TRadioButton;
    NEPlafond: THNumEdit;
    CBCommission: THValComboBox;
    TJournalVente: THLabel;
    CMJournalVente: THCritMaskEdit;
    CMContexte: THCritMaskEdit;
    GB_Stocks: TGroupBox;
    HLabel2: THLabel;
    HLabel5: THLabel;
    DateDebutExercice: THCritMaskEdit;
    NbMois: THNumEdit;
    GroupBox2: TGroupBox;
    HLabel10: THLabel;
    CB_Dossier: THValComboBox;
    TabSheet8: TTabSheet;
    GroupBox3: TGroupBox;
    HLabel11: THLabel;
    Button2: TButton;
    CB_FicTra: TEdit;
    bNewDossier: TBitBtn;
    Edit1: TEdit;
    bSupDossier: TBitBtn;
    LBTri: TListBox;
    procedure FormShow(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure bNewDossierClick(Sender: TObject);
    procedure bSupDossierClick(Sender: TObject);
    procedure StringGrid1Click(Sender: TObject);
    procedure FormClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure StringGrid1DblClick(Sender: TObject);
    procedure StringGrid1SelectCell(Sender: TObject; ACol, ARow: Integer; var CanSelect: Boolean);
    procedure BAssociationChampsClick(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure B_RepertoireClick(Sender: TObject);
    procedure Edit1Enter(Sender: TObject);
    procedure Edit1Change(Sender: TObject);
    procedure Edit1Exit(Sender: TObject);
    procedure CB_DossierChange(Sender: TObject);
    procedure CB_RepertoireChange(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

  procedure Assist_Import ;

var
    Tob_Detables, Tob_Dechamp : Tob;
    i_NumEcran : integer;
    st_Chemin, st_Suffixe : string;

implementation

{$R *.DFM}

Procedure Assist_Import ;
var
   Fo_Assist : TFAssistimport;
Begin
     Fo_Assist := TFAssistimport.Create (Application);
     Try
         Fo_Assist.ShowModal;
     Finally
         Fo_Assist.free;
     End;
end;

procedure TFAssistimport.FormShow(Sender: TObject);
var
    i_ind1 : integer;
    Onglet : TTabSheet;
    St_NomOnglet, stDefaut : String;
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

    StringGrid1.ColWidths[0] := 125;
    StringGrid1.ColWidths[1] := 245;

    st_Chemin := ExtractFilePath(ParamStr(0));
    if FileExists(st_Chemin + 'ListeRepertoire') then
    begin
        Ficliste := TIniFile.Create(st_Chemin + 'ListeRepertoire');
        Ficliste.ReadSectionValues('General', CB_Repertoire.Items);
        FicListe.Free;
        for i_ind1 := 0 to CB_Repertoire.Items.Count - 1 do
            begin
            stDefaut := CB_Repertoire.Items.Strings[i_ind1];
            stDefaut := Copy(stDefaut, Pos('=', stDefaut) + 1, 255);
            CB_Repertoire.Items.Strings[i_ind1] := stDefaut;
            end;
{    end
    else
    begin
    CB_Repertoire.Values.Add(stDefaut);
    CB_Repertoire.Items.Add(stDefaut);}
    end;
    CB_Repertoire.ItemIndex := 0;

    if CB_Repertoire.Text <> '' then
        begin
        st_Chemin := CB_Repertoire.Text;
        if not FileExists(st_Chemin + 'ListeFic') then
        begin
            Ficliste := TIniFile.Create(st_Chemin + 'ListeFic');
            Ficliste.WriteString ('DossierDefaut', 'Defaut', 'Négoce');
            Ficliste.WriteString ('DossierActif', 'Actif', 'Dossier0');
            Ficliste.WriteString ('Dossier', 'Dossier0', 'Négoce');
            Ficliste.WriteString ('Dossier', 'Dossier1', 'Fichier .tra');
            Ficliste.Free;
        end;
        Ficliste := TIniFile.Create(st_Chemin + 'ListeFic');
        Ficliste.ReadSectionValues('Dossier', CB_Dossier.Items);
        for i_ind1 := 0 to CB_Dossier.Items.Count - 1 do
            begin
            stDefaut := CB_Dossier.Items.Strings[i_ind1];
            stDefaut := Copy(stDefaut, Pos('=', stDefaut) + 1, 255);
            CB_Dossier.Items.Strings[i_ind1] := stDefaut;
            end;
        stDefaut := FicListe.ReadString('DossierDefaut', 'Defaut', '');
        CB_Dossier.ItemIndex := CB_Dossier.Items.IndexOf(stDefaut);
        FicListe.Free;
        end;
end;

procedure TFAssistimport.bSuivantClick(Sender: TObject);
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
    0 :
        begin
            // bFin.Enabled := True;
            Ficliste := TIniFile.Create(st_Chemin + 'ListeFic');
            Ficliste.EraseSection ('General' + st_Suffixe);
            if CB_inibase.checked then
                Ficliste.WriteString ('General' + st_Suffixe, 'DelFic', 'O')
                else
                Ficliste.WriteString ('General' + st_Suffixe, 'DelFic', 'N');
            if RB_Creation.checked then
                Ficliste.WriteString ('General' + st_Suffixe, 'ModeCreation', 'O')
                else if RB_Maj.checked then
                Ficliste.WriteString ('General' + st_Suffixe, 'ModeCreation', 'M')
                else
                Ficliste.WriteString ('General' + st_Suffixe, 'ModeCreation', 'P');
            if RB_clientC.checked then
                Ficliste.WriteString ('General' + st_Suffixe, 'Client', '"C"')
                else
                Ficliste.WriteString ('General' + st_Suffixe, 'Client', '"9"');
            if RB_FourniF.checked then
                Ficliste.WriteString ('General' + st_Suffixe, 'Fournisseur', '"F"')
                else
                Ficliste.WriteString ('General' + st_Suffixe, 'Fournisseur', '"0"');
            if RBMonnaieFRF.checked then
                Ficliste.WriteString ('General' + st_Suffixe, 'Monnaie', 'F')
                else
                Ficliste.WriteString ('General' + st_Suffixe, 'Monnaie', 'E');
            Ficliste.WriteString ('General' + st_Suffixe, 'Plafond', FloatToStr(NEPlafond.Value));
            Ficliste.WriteString ('General' + st_Suffixe, 'CommissionText', CBCommission.Text);
            Ficliste.WriteString ('General' + st_Suffixe, 'CommissionValue', CBCommission.Value);
            Ficliste.WriteString ('General' + st_Suffixe, 'NoStatClient', IntToStr(SE_Client.Value));
            Ficliste.WriteString ('General' + st_Suffixe, 'QualifBarreText', QualifCodeBarre.Text);
            Ficliste.WriteString ('General' + st_Suffixe, 'QualifBarreVal', QualifCodeBarre.Value);
            Ficliste.WriteString ('General' + st_Suffixe, 'JournalVente', CMJournalVente.Text);
            Ficliste.WriteString ('General' + st_Suffixe, 'ColClient', CollectifCli.Text);
            Ficliste.WriteString ('General' + st_Suffixe, 'ColFourni', CollectifFou.Text);
            Ficliste.WriteString ('General' + st_Suffixe, 'NoStatClient', IntToStr(SE_Client.Value));
            Ficliste.WriteString ('General' + st_Suffixe, 'NoStatArticle', IntToStr(SE_Article.Value));
            Ficliste.WriteString ('General' + st_Suffixe, 'DebutExercice', DateDebutExercice.Text);
            Ficliste.WriteString ('General' + st_Suffixe, 'NbMois', NbMois.Text);

            Ts_listeFic := TstringList.Create;
            i_ind1 := 0;
            Ts_listeFic.Text := ''; // Réinitialisation de la TString
            Ficliste.ReadSectionValues ('Fichier' + st_Suffixe, Ts_listeFic);
            Tob_Detables := TOB.Create('DETABLES', nil, -1);
            Tob_Detables.LoadDetailDB('DETABLES','','',Nil,FALSE);
            if Ts_listeFic.Count = 0 then
                begin
                StringGrid1.RowCount := Tob_Detables.Detail.Count;
                Tob_Detables.PutStringsDetail(StringGrid1.Cols[0],'DT_NOMTABLE');
                end
                else
                begin
                for i_ind1 := 0 to Tob_Detables.Detail.Count - 1 do
                    begin
                    st_enreg := Tob_Detables.Detail[i_ind1].GetValue('DT_NOMTABLE');
                    if Pos(st_enreg, Ts_ListeFic.Text) = 0 then
                        Ts_ListeFic.Add(st_enreg + '=;;;;');
                    end;
                LBTri.Items.Text := Ts_ListeFic.Text;
                Ts_ListeFic.Text := LBTri.Items.Text;
                StringGrid1.RowCount := Ts_ListeFic.Count;
                i_ind1 := 0;
                while i_ind1 <= Ts_listeFic.Count - 1 do
                    begin
                    st_enreg := Ts_listeFic.Strings [i_ind1];
                    StringGrid1.Cells[0, i_ind1] := Copy(st_enreg, 0, Pos('=', st_enreg) - 1);
                    st_enreg := Copy(st_enreg, Pos('=', st_enreg) + 1, 255);
                    i_ind2 := 1;
                    while Pos(';', st_enreg) <> 0 do
                    begin
                        StringGrid1.Cells[i_ind2, i_ind1] := Copy(st_enreg, 0, Pos(';', st_enreg) - 1);
                        st_enreg := Copy(st_enreg, Pos(';', st_enreg) + 1, 255);
                        i_ind2 := i_ind2 + 1;
                    end;
                    StringGrid1.Cells[i_ind2, i_ind1] := st_enreg;
                    i_ind1 := i_ind1 + 1;
                    end;
                end;
            Tob_Detables.Free;
            Ts_listeFic.Free;
            FicListe.Free;
        end;
    1 :
        begin
            Ficliste := TIniFile.Create(st_Chemin + 'ListeFic');
            if not FicListe.ValueExists('General' + st_Suffixe, 'BeforeAction') then
                begin
                ListRecap.Items.Add ('---> ' + GB_CodeBarre.Caption);
                ListRecap.Items.Add (TQualifCodeBarre.Caption + ' : ' + QualifCodeBarre.Text);
                ListRecap.Items.Add ('');
                ListRecap.Items.Add ('---> ' + GB_compteclient.Caption);
                if RB_clientC.Checked then
                    st_chaine := ' : "C"'
                    else
                    st_chaine := ' : "9"';
                ListRecap.Items.Add (TDebutCompteCli.Caption + st_chaine);
                ListRecap.Items.Add (TCollectifCli.Caption + ' : ' + CollectifCli.Text);
                ListRecap.Items.Add (TJournalVente.Caption + ' : ' + CMJournalVente.Text);
                ListRecap.Items.Add ('');
                ListRecap.Items.Add ('---> ' + GB_comptefour.Caption);
                if RB_FourniF.Checked then
                    st_chaine := ' : "F"'
                    else
                    st_chaine := ' : "0"';
                ListRecap.Items.Add (TDebutCompteFou.Caption + st_chaine);
                ListRecap.Items.Add (TCollectifFou.Caption + ' : ' + CollectifFou.Text);
                ListRecap.Items.Add ('');
                ListRecap.Items.Add ('---> ' + GB_Monnaie.Caption);
                if RBMonnaieFRF.Checked then
                    st_chaine := ' : Francs'
                    else
                    st_chaine := ' : Euros';
                ListRecap.Items.Add (TMonnaieTenue.Caption + st_chaine);
                ListRecap.Items.Add (TPlafond.Caption + ' : ' + FloatToStr(NEPlafond.Value));
                ListRecap.Items.Add (TCommission.Caption + ' : ' + CBCommission.Text);

                RestorePage ;
                Onglet := NextPage ;
                if Onglet = nil then
                    P.SelectNextPage(True)
                    else
                    begin
                    P.ActivePage := Onglet ;
                    PChange(nil) ;
                    end ;
                end;
            FicListe.Free;
        end;
    2 :
        begin
            Ficliste := TIniFile.Create(st_Chemin + 'ListeFic');
            Ficliste.EraseSection ('Fichier' + st_Suffixe);
            i_ind1 := 0;
            while i_ind1 <= StringGrid1.RowCount do
            begin
                    if StringGrid1.Cells[0, i_ind1] <> '' then
                    begin
                        st_ident := StringGrid1.Cells[0, i_ind1];
                        st_valeur := '';
                        for i_ind2 := 1 to 5 do
                        begin
                            if i_ind2 <> 5 then
                                st_valeur := st_valeur + StringGrid1.Cells[i_ind2, i_ind1] + ';'
                            else
                                st_valeur := st_valeur + StringGrid1.Cells[i_ind2, i_ind1];
                        end;
                        FicListe.WriteString('Fichier' + st_Suffixe, st_ident, st_valeur);
                    end;
                i_ind1 := i_ind1 + 1;
            end;
            FicListe.Free;
            if not V_PGI.Superviseur then
                begin
                RestorePage ;
                Onglet := NextPage ;
                if Onglet = nil then
                    P.SelectNextPage(True)
                    else
                    begin
                    P.ActivePage := Onglet ;
                    PChange(nil) ;
                    end ;
                end;
        end;
    3 :
        begin
        Ficliste := TIniFile.Create(st_Chemin + 'ListeFic');
        if FicListe.ValueExists('General' + st_Suffixe, 'BeforeAction') then
            begin
            Ts_listeFic := TstringList.Create;
            i_ind1 := 0;
            Ts_listeFic.Text := ''; // Réinitialisation de la TString
            Ficliste.ReadSectionValues ('Fichier' + st_Suffixe, Ts_listeFic);
            Tob_Detables := TOB.Create('DETABLES', nil, -1);
            Tob_Detables.LoadDetailDB('DETABLES','','',Nil,FALSE);
            if Ts_listeFic.Count = 0 then
                begin
                StringGrid1.RowCount := Tob_Detables.Detail.Count;
                Tob_Detables.PutStringsDetail(StringGrid1.Cols[0],'DT_NOMTABLE');
                end
                else
                begin
                for i_ind1 := 0 to Tob_Detables.Detail.Count - 1 do
                    begin
                    st_enreg := Tob_Detables.Detail[i_ind1].GetValue('DT_NOMTABLE');
                    if Pos(st_enreg, Ts_ListeFic.Text) = 0 then
                        Ts_ListeFic.Add(st_enreg + '=;;;;');
                    end;
                LBTri.Items.Text := Ts_ListeFic.Text;
                Ts_ListeFic.Text := LBTri.Items.Text;
                StringGrid1.RowCount := Ts_ListeFic.Count;
                i_ind1 := 0;
                while i_ind1 <= Ts_listeFic.Count - 1 do
                    begin
                    st_enreg := Ts_listeFic.Strings [i_ind1];
                    StringGrid1.Cells[0, i_ind1] := Copy(st_enreg, 0, Pos('=', st_enreg) - 1);
                    st_enreg := Copy(st_enreg, Pos('=', st_enreg) + 1, 255);
                    i_ind2 := 1;
                    while Pos(';', st_enreg) <> 0 do
                    begin
                        StringGrid1.Cells[i_ind2, i_ind1] := Copy(st_enreg, 0, Pos(';', st_enreg) - 1);
                        st_enreg := Copy(st_enreg, Pos(';', st_enreg) + 1, 255);
                        i_ind2 := i_ind2 + 1;
                    end;
                    StringGrid1.Cells[i_ind2, i_ind1] := st_enreg;
                    i_ind1 := i_ind1 + 1;
                    end;
                end;
            Tob_Detables.Free;
            Ts_listeFic.Free;
            repeat
                RestorePage ;
                Onglet := NextPage ;
                if Onglet = nil then
                    P.SelectNextPage(True)
                    else
                    begin
                    P.ActivePage := Onglet ;
                    PChange(nil) ;
                    end ;
                st_NomOnglet := Onglet.Name;
                i_NumEcran := strtoint (Copy (st_NomOnglet, length (st_NomOnglet), 1)) - 1;
            until (i_NumEcran = 1) or
                  ((i_NumEcran = 2) and V_PGI.Superviseur);
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
            st_Suffixe := '';
            if CB_Dossier.ItemIndex >= 0 then
                if CB_Dossier.Items.Strings[CB_Dossier.ItemIndex] <> '' then
                    st_Suffixe := '_Dossier' + IntToStr(CB_Dossier.ItemIndex);
            st_Chemin := CB_Repertoire.Text;
            if FileExists(st_Chemin + 'ListeFic') then
            begin
                Ficliste := TIniFile.Create(st_Chemin + 'ListeFic');
                if CB_Dossier.Tag <> 0 then
                    begin
                    CB_Dossier.ItemIndex := CB_Dossier.Tag;
                    CB_Dossier.Tag := 0;
                    end;
                if CB_Dossier.Text = 'Fichier .tra' then
                    Ficliste.WriteString ('General' + st_Suffixe, 'BeforeAction', 'O');
                Ficliste.WriteString ('DossierActif', 'Actif', 'Dossier' + IntToStr(CB_Dossier.ItemIndex));
                for i_ind1 := 0 to CB_Dossier.Items.Count - 1 do
                    begin
                    st_enreg := CB_Dossier.Items.Strings[i_ind1];
                    Ficliste.WriteString ('Dossier', 'Dossier'+IntToStr(i_ind1), st_enreg);
                    end;
                if FicListe.ValueExists('General' + st_Suffixe, 'BeforeAction') then
                    begin
                    st_enreg := Ficliste.ReadString ('General' + st_Suffixe, 'BeforeAction', '');
                    end
                    else
                    begin
                    st_enreg := Ficliste.ReadString ('General' + st_Suffixe, 'DelFic', '');
                    if st_enreg = 'O' then
                        CB_inibase.checked := True
                        else
                        CB_inibase.checked := False;
                    st_enreg := Ficliste.ReadString ('General' + st_Suffixe, 'ModeCreation', '');
                    if st_enreg = 'O' then
                        RB_Creation.checked := True
                        else if st_enreg = 'M' then
                        RB_Maj.checked := True
                        else
                        RB_MajP.checked := True;
                    st_enreg := Ficliste.ReadString ('General' + st_Suffixe, 'Client', '');
                    if st_enreg = 'C' then
                        RB_clientC.checked := True
                        else
                        RB_client9.checked := True;
                    st_enreg := Ficliste.ReadString ('General' + st_Suffixe, 'Fournisseur', '');
                    if st_enreg = 'F' then
                        RB_FourniF.checked := True
                        else
                        RB_Fourni0.checked := True;
                    st_enreg := Ficliste.ReadString ('General' + st_Suffixe, 'Monnaie', '');
                    if st_enreg = 'F' then
                        RBMonnaieFRF.checked := True
                        else
                        RBMonnaieEURO.checked := True;
                    CBCommission.Text := Ficliste.ReadString ('General' + st_Suffixe, 'CommissionText', '');
                    CBCommission.Value := Ficliste.ReadString ('General' + st_Suffixe, 'CommissionValue', '');
                    NEPlafond.Value := Ficliste.ReadFloat ('General' + st_Suffixe, 'Plafond', 0.0);
                    QualifCodeBarre.Text := Ficliste.ReadString ('General' + st_Suffixe, 'QualifBarreText', '');
                    QualifCodeBarre.Value := Ficliste.ReadString ('General' + st_Suffixe, 'QualifBarreVal', '');
                    CMJournalVente.Text := Ficliste.ReadString ('General' + st_Suffixe, 'JournalVente', '');
                    CollectifCli.Text := Ficliste.ReadString ('General' + st_Suffixe, 'ColClient', '');
                    CollectifFou.Text := Ficliste.ReadString ('General' + st_Suffixe, 'ColFourni', '');
                    SE_Client.Value := Ficliste.ReadInteger ('General' + st_Suffixe, 'NoStatClient', 0);
                    SE_Article.Value := Ficliste.ReadInteger ('General' + st_Suffixe, 'NoStatArticle', 0);
                    end;

                Ts_listeFic := TstringList.Create;
                i_ind1 := 0;
                Ts_listeFic.Text := ''; // Réinitialisation de la TString
                Ficliste.ReadSectionValues ('Fichier' + st_Suffixe, Ts_listeFic);
                if Ts_listeFic.Count > 0 then
                begin
                    while i_ind1 <= Ts_listeFic.Count - 1 do
                    begin
                        st_enreg := Ts_listeFic.Strings [i_ind1];
                        StringGrid1.Cells[0, i_ind1] := Copy(st_enreg, 0, Pos('=', st_enreg) - 1);
                        st_enreg := Copy(st_enreg, Pos('=', st_enreg) + 1, 255);
                        i_ind2 := 1;
                        while Pos(';', st_enreg) <> 0 do
                        begin
                            StringGrid1.Cells[i_ind2, i_ind1] := Copy(st_enreg, 0, Pos(';', st_enreg) - 1);
                            st_enreg := Copy(st_enreg, Pos(';', st_enreg) + 1, 255);
                            i_ind2 := i_ind2 + 1;
                        end;
                        StringGrid1.Cells[i_ind2, i_ind1] := st_enreg;
                        i_ind1 := i_ind1 + 1;
                    end;
                end
                else
                begin
                    Tob_Detables := TOB.Create('DETABLES', nil, -1);
                    Tob_Detables.LoadDetailDB('DETABLES','','',Nil,FALSE);
                    StringGrid1.RowCount := Tob_Detables.Detail.Count;
                    Tob_Detables.PutStringsDetail(StringGrid1.Cols[0],'DT_NOMTABLE');
                    Tob_Detables.Free;
                end;
                Ts_listeFic.Free;
                Ficliste.Free;
            end;
        end;
    7 :
        begin
        Ficliste := TIniFile.Create(st_Chemin + 'ListeFic');
        ListRecap.Items.Clear;
        ListRecap.Items.Add ('---> ' + GB_inibase.Caption);
        if CB_inibase.checked then
            st_chaine := ' : Oui'
            else
            st_chaine := ' : Non';
        ListRecap.Items.Add (CB_inibase.Caption + st_chaine);
        if RB_Creation.Checked Then
            ListRecap.Items.Add (RB_Creation.Caption)
            else
            ListRecap.Items.Add (RB_Maj.Caption);
        ListRecap.Items.Add ('');
        ListRecap.Items.Add ('---> Fichier traité');
        ListRecap.Items.Add (CB_FicTra.Text + '.tra');
        ListRecap.Items.Add ('');
        FicListe.Free;
        end;
    end;
    Onglet := P.ActivePage;
    st_NomOnglet := Onglet.Name;
    i_NumEcran := strtoint (Copy (st_NomOnglet, length (st_NomOnglet), 1)) - 1;
    //i_NumEcran := i_NumEcran + 1;
    Image.ImageIndex := i_NumEcran;
    if (bSuivant.Enabled) then
    begin
        bFin.Enabled := False;
    end
    else
    begin
        bFin.Enabled := True;
    end;
end;

procedure TFAssistimport.bPrecedentClick(Sender: TObject);
Var
    Onglet : TTabSheet;
    St_NomOnglet : String;
    Ficliste : TIniFile;
begin
    inherited;
    Onglet := P.ActivePage;
    st_NomOnglet := Onglet.Name;
    i_NumEcran := strtoint (Copy (st_NomOnglet, length (st_NomOnglet), 1)) - 1;
    case i_NumEcran of
    0 :
        begin
        Ficliste := TIniFile.Create(st_Chemin + 'ListeFic');
        if FicListe.ValueExists('General' + st_Suffixe, 'BeforeAction') then
            begin
            repeat
                RestorePage ;
                Onglet := PreviousPage ;
                if Onglet = nil then
                    P.SelectNextPage(True)
                    else
                    begin
                    P.ActivePage := Onglet ;
                    PChange(nil) ;
                    end ;
                st_NomOnglet := Onglet.Name;
                i_NumEcran := strtoint (Copy (st_NomOnglet, length (st_NomOnglet), 1)) - 1;
            until (i_NumEcran = 3);
            end;
        FicListe.Free;
        end;
    1 :
        begin
            if not V_PGI.Superviseur then
            begin
            RestorePage ;
            Onglet := PreviousPage ;
            if Onglet = nil then
                P.SelectNextPage(True)
                else
                begin
                P.ActivePage := Onglet ;
                PChange(nil) ;
                end ;
            end;
        end;
    2 :
        begin
        Ficliste := TIniFile.Create(st_Chemin + 'ListeFic');
        if (FicListe.ValueExists('General' + st_Suffixe, 'BeforeAction')) and (not V_PGI.Superviseur) then
            begin
            repeat
                RestorePage ;
                Onglet := PreviousPage ;
                if Onglet = nil then
                    P.SelectNextPage(True)
                    else
                    begin
                    P.ActivePage := Onglet ;
                    PChange(nil) ;
                    end ;
                st_NomOnglet := Onglet.Name;
                i_NumEcran := strtoint (Copy (st_NomOnglet, length (st_NomOnglet), 1)) - 1;
            until (i_NumEcran = 3);
            end;
        FicListe.Free;
        end;
    7 :
        begin
        Ficliste := TIniFile.Create(st_Chemin + 'ListeFic');
        if not FicListe.ValueExists('General' + st_Suffixe, 'BeforeAction') then
            begin
            RestorePage ;
            Onglet := PreviousPage ;
            if Onglet = nil then
                P.SelectNextPage(True)
                else
                begin
                P.ActivePage := Onglet ;
                PChange(nil) ;
                end ;
            end;
        FicListe.Free;
        end;
    end;
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

procedure TFAssistimport.Button1Click(Sender: TObject);
begin
    inherited;
    OpenDialog1.InitialDir := CB_Repertoire.Text;
    OpenDialog1.FileName := '*.des';
    OpenDialog1.Filter := 'Descriptifs PGI (*.des)|*.des';
    if OpenDialog1.Execute then
        if OpenDialog1.FileName <> '' then
            StringGrid1.Cells[StringGrid1.Col, StringGrid1.Row] :=
                Copy(OpenDialog1.Filename, 0, Pos('.', OpenDialog1.Filename) - 1);
//    else
  //      StringGrid1.Cells[StringGrid1.Col, StringGrid1.Row] := '';
    Button1.Visible := False;
end;

procedure TFAssistimport.Button2Click(Sender: TObject);
begin
    inherited;
    OpenDialog1.InitialDir := CB_Repertoire.Text;
    OpenDialog1.FileName := '*.tra';
    OpenDialog1.Filter := 'Fichiers transfert (*.tra)|*.tra';
    if OpenDialog1.Execute then
        if OpenDialog1.FileName <> '' then
            CB_FicTra.Text := Copy(OpenDialog1.Filename, 0, Pos('.', OpenDialog1.Filename) - 1);
    Button2.Visible := False;
end;

procedure TFAssistimport.bNewDossierClick(Sender: TObject);
begin
  inherited;
  CB_Dossier.Visible := False;
  Edit1.Visible := True;
  Edit1.SetFocus;
  bNewDossier.Visible := False;
end;

procedure TFAssistimport.bSupDossierClick(Sender: TObject);
var
    Ficliste : TIniFile;
begin
Ficliste := TIniFile.Create(st_Chemin + 'ListeFic');
if FicListe.ValueExists('Dossier', 'Dossier' + IntToStr(CB_Dossier.ItemIndex)) then
    Ficliste.DeleteKey('Dossier', 'Dossier' + IntToStr(CB_Dossier.ItemIndex));
if FicListe.SectionExists('General' + st_Suffixe) then
    Ficliste.EraseSection ('General' + st_Suffixe);
if FicListe.SectionExists('Fichier' + st_Suffixe) then
    Ficliste.EraseSection ('Fichier' + st_Suffixe);
FicListe.Free;
CB_Dossier.Items.Delete(CB_Dossier.ItemIndex);
CB_Dossier.ItemIndex := CB_Dossier.Items.Count - 1;
CB_DossierChange(Sender);
end;

procedure TFAssistimport.StringGrid1Click(Sender: TObject);
var
    Sel_Rect : TRect;

begin
  inherited;
//    Button1.Visible := False;
    Sel_Rect := StringGrid1.CellRect(StringGrid1.Col, StringGrid1.Row);
    Button1.Top := Sel_Rect.Top + StringGrid1.Top + 2;
    Button1.Left := Sel_Rect.Right - Button1.Width + StringGrid1.Left + 2;
    Button1.Height := Sel_Rect.Bottom - Sel_Rect.Top;
    Button1.Visible := True;
end;

procedure TFAssistimport.FormClick(Sender: TObject);
begin
  inherited;
    Button1.Visible := False;
end;

procedure TFAssistimport.FormCreate(Sender: TObject);
begin
  inherited;
    FAssocImport := TFAssocImport.Create(FAssocImport);
end;

procedure TFAssistimport.StringGrid1DblClick(Sender: TObject);
begin
  inherited;
    Button1Click(Sender);
end;

procedure TFAssistimport.StringGrid1SelectCell(Sender: TObject; ACol,
  ARow: Integer; var CanSelect: Boolean);
var
    Sel_Rect : TRect;

begin
  inherited;
    Sel_Rect := StringGrid1.CellRect(StringGrid1.Col, StringGrid1.Row);
    Button1.Top := Sel_Rect.Top + StringGrid1.Top + 2;
    Button1.Left := Sel_Rect.Right - Button1.Width + StringGrid1.Left + 2;
    Button1.Height := Sel_Rect.Bottom - Sel_Rect.Top;
    Button1.Visible := True;
end;

procedure TFAssistimport.BAssociationChampsClick(Sender: TObject);
var
    i_ind1, i_ind2 : integer;
begin
    inherited;

    i_ind2 := 0;
    for i_ind1 := 0 to stringgrid1.rowcount do
        if stringgrid1.cells[1, i_ind1] <> '' then
        begin
            fassocimport.stringgrid1.rows[i_ind2] := stringgrid1.rows[i_ind1];
            i_ind2 := i_ind2 + 1;
        end;
    fassocimport.stringgrid1.width := stringgrid1.width;
    fassocimport.stringgrid1.colwidths[0] := stringgrid1.colwidths[0];
    fassocimport.stringgrid1.colwidths[1] := stringgrid1.colwidths[1];
    fassocimport.stringgrid1.colwidths[2] := stringgrid1.colwidths[2];
    fassocimport.stringgrid1.colwidths[3] := stringgrid1.colwidths[3];
    fassocimport.stringgrid1.colwidths[4] := stringgrid1.colwidths[4];
    fassocimport.stringgrid1.colwidths[5] := stringgrid1.colwidths[5];
    fassocimport.WorkingPath.Text := st_Chemin;
    fassocimport.showmodal;
end;

procedure TFAssistimport.bFinClick(Sender: TObject);
var FicListe : TIniFile;
begin
    inherited;

    Ficliste := TIniFile.Create(st_Chemin + 'ListeFic');
    if FicListe.ValueExists('General' + st_Suffixe, 'BeforeAction') then
        begin
        if Genere_PGI(CB_FicTra.Text) then
            RecuperationNeg(st_Chemin);
        end
        else
        RecuperationNeg(st_Chemin);
    FicListe.Free;
    Close ;
end;


procedure TFAssistimport.B_RepertoireClick(Sender: TObject);
var
   Directory, stDefaut : string;
   HelpCtx: Longint;
   Ficliste : TIniFile;
   i_ind1 : integer;
begin
  inherited;
  if SelectDirectory(Directory, [sdAllowCreate, sdPrompt], HelpCtx) then
     Begin
     CB_Repertoire.Value := Directory + '\';
     CB_Repertoire.Values.Add(Directory + '\');
     CB_Repertoire.Items.Add(Directory + '\');
     CB_Repertoire.Text := Directory + '\';
     SetCurrentDir(Directory);
     CB_RepertoireChange(Sender);
     end;
end;

procedure TFAssistimport.CB_RepertoireChange(Sender: TObject);
var
   stDefaut : string;
   HelpCtx: Longint;
   Ficliste : TIniFile;
   i_ind1 : integer;
begin
if CB_Repertoire.Text <> '' then
    begin
    CB_Dossier.Items.Clear;
    stDefaut := CB_Repertoire.Text;
    if FileExists(stDefaut + 'ListeFic') then
    begin
        Ficliste := TIniFile.Create(stDefaut + 'ListeFic');
        Ficliste.ReadSectionValues('Dossier', CB_Dossier.Items);
        for i_ind1 := 0 to CB_Dossier.Items.Count - 1 do
            begin
            stDefaut := CB_Dossier.Items.Strings[i_ind1];
            stDefaut := Copy(stDefaut, Pos('=', stDefaut) + 1, 255);
            CB_Dossier.Items.Strings[i_ind1] := stDefaut;
            end;
        stDefaut := FicListe.ReadString('DossierDefaut', 'Defaut', '');
        CB_Dossier.ItemIndex := CB_Dossier.Items.IndexOf(stDefaut);
        FicListe.Free;
    end
    else
    begin
    CB_Dossier.Items.Add('Négoce');
    CB_Dossier.Items.Add('Fichier .tra');
    CB_Dossier.ItemIndex := 0;
    end;
    end;
end;

procedure TFAssistimport.Edit1Enter(Sender: TObject);
var ind1 : integer;
begin
ind1 := CB_Dossier.Items.Add('Nouveau');
CB_Dossier.Tag := ind1;
end;

procedure TFAssistimport.Edit1Change(Sender: TObject);
begin
CB_Dossier.Items.Strings[CB_Dossier.Tag] := Edit1.Text;
CB_Dossier.ItemIndex := CB_Dossier.Tag;
end;

procedure TFAssistimport.Edit1Exit(Sender: TObject);
begin
Edit1.Visible := False;
CB_Dossier.Visible := True;
CB_Dossier.ItemIndex := CB_Dossier.Tag;
end;

procedure TFAssistimport.CB_DossierChange(Sender: TObject);
begin
st_Suffixe := '';
if CB_Dossier.ItemIndex >= 0 then
    if CB_Dossier.Items.Strings[CB_Dossier.ItemIndex] <> '' then
        st_Suffixe := '_Dossier' + IntToStr(CB_Dossier.ItemIndex);
end;

end.
