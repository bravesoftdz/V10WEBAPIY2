unit AssocImport;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, StdCtrls, ExtCtrls, Buttons, UTob, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} UiUtil, Hctrls,
  ComCtrls, IniFiles, AssistUtil;

type
  TFAssocImport = class(TForm)
    Edit1: TEdit;
    StringGridFichier: TStringGrid;
    StringGridTable: TStringGrid;
    StringGridAsso: TStringGrid;
    StringGrid1: TStringGrid;
    Edit2: TEdit;
    OpenDialog1: TOpenDialog;
    StringGridAssoStruct: TStringGrid;
    Panel1: TPanel;
    ListeTable: TListBox;
    ListeAsso: TListBox;
    ListeFichier: TListBox;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    BitBtn3: TBitBtn;
    BitBtn4: TBitBtn;
    RadioGroup1: TRadioGroup;
    BitBtn5: TBitBtn;
    ListeAssoStruct: TListBox;
    RadioGroup2: TRadioGroup;
    TabControl1: TTabControl;
    ListeAssoCh: TListBox;
    StringGridAssoFic: TStringGrid;
    RadioGroup3: TRadioGroup;
    StringGrid2: TStringGrid;
    BitBtn6: TBitBtn;
    Edit3: TEdit;
    BitBtn7: TBitBtn;
    BitBtn8: TBitBtn;
    StringGrid3: TStringGrid;
    Edit4: TEdit;
    Label1: TLabel;
    BitBtn9: TBitBtn;
    RadioGroup4: TRadioGroup;
    BitBtn10: TBitBtn;
    BitBtn11: TBitBtn;
    WorkingPath: TEdit;
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
    procedure BitBtn5Click(Sender: TObject);
    procedure ListeFichierDblClick(Sender: TObject);
    procedure RadioGroup1Click(Sender: TObject);
    procedure ListeTableDblClick(Sender: TObject);
    procedure ListeDictFichier;
    procedure ListeDetailAsso;
    procedure StringGrid1DblClick(Sender: TObject);
    procedure RadioGroup2Click(Sender: TObject);
    procedure TabControl1Change(Sender: TObject);
    procedure RadioGroup3Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BitBtn6Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure ListeAssoClick(Sender: TObject);
    procedure ListeAssoStructClick(Sender: TObject);
    procedure BitBtn7Click(Sender: TObject);
    procedure BitBtn8Click(Sender: TObject);
    procedure ListeAssoChClick(Sender: TObject);
    procedure ListeTableEnter(Sender: TObject);
    procedure ListeFichierKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure BitBtn9Click(Sender: TObject);
    procedure RadioGroup4Click(Sender: TObject);
    procedure BitBtn10Click(Sender: TObject);
    procedure BitBtn11Click(Sender: TObject);
  private
    { Déclarations privées }
    TobTable : TOB;
    procedure Charge_Dictionnaire;
    procedure Enregistre_Inifile(i_col,i_row : integer);
    function  VerifCorrespondance (i_nb_champs : integer; st_champs, st_racine : string; var i_nrow : integer) : boolean;
    procedure Charge_Section(FicIni : TIniFile; st_racine : string;
                             i_col : integer; var i_ligne : integer);
    procedure Maj_Edit3;
    procedure ListeTableEntree(Sender: TObject; var Key: Word);
  public
    { Déclarations publiques }
  end;

var
  FAssocImport: TFAssocImport;
  stDossier : string;

implementation

uses SaisieTexte, SaisieFormule, SaisieCondition{, AssistImport};

{$R *.DFM}
var
    i_AssoRow, i_AssoCol, i_TabIndexPrec, i_RowPrec : integer;
    st_Chemin, st_racine : string;
    st_ListeTable_Rech, st_ListeAsso_Rech, st_ListeFichier_Rech : string;

procedure TFAssocImport.Charge_Dictionnaire;
var
    Tob1, Tob2 : TOB;
    Query1 : TQuery;
    Fichier: textfile;
    st_enreg, st_trav1, st_trav2, st_file, st_maj : string;
    i_ind1, i_ind2, i_ind3, i_ind4, i_indtext, i_nrow : integer;
    i_nb_ch_text, i_ligne_stringgrid : integer;
    b_Trouve : boolean;
    FicIni : TIniFile;
    ParamGenCle : TStrings;
    ParamGenData : TStrings;

begin
    if Edit2.Text = '' then Exit;
    //  Chargement de la version precedente eventuelle
    i_ind1 := 0;
    i_ind2 := 0;
    i_ind3 := 0;
    i_AssoRow := 0;
    i_AssoCol := 0;
    st_Chemin := WorkingPath.Text;
    st_file := st_Chemin + Edit1.Text;
    if not FileExists (st_file + '.ini') then
    begin
        if FileExists (st_file + '.def') then
        begin
            CopyFile (PChar (st_file + '.def'), PChar (st_file + '.ini'), False);
        end;
    end;
    for i_ind1 := 0 to StringGridFichier.RowCount - 1 do
    begin
        if StringGridFichier.cells[0, i_ind1] = '' then break;
        StringGridFichier.Rows[i_ind1].Clear;
    end;
//  Ouverture dictionnaire fichier texte
    AssignFile(Fichier, Edit2.Text + '.des');
    Reset(Fichier);
//  Chargement de la liste des champs du fichier texte
    readln(Fichier, st_enreg);
    st_racine := Copy(st_enreg, 0, 3);
//  au cas ou, on met la racine dans stringgrid3
    StringGrid3.Cells[TabControl1.TabIndex + 1, StringGrid1.Row] := st_racine;
    i_ind1 := 0;
    i_ind3 := 0;
    while st_enreg <> '' do
    begin
        i_ind2 := 0;
        st_trav1 := Trim(Copy(st_enreg, 1, Pos(';',st_enreg) - 1));
        st_trav2 := Trim(Copy(st_enreg, Pos(';',st_enreg)+1, 255));
        StringGridFichier.Cells[0, i_ind1] := st_trav1;
        st_enreg := st_trav2;
        st_trav1 := Trim(Copy(st_enreg, 1, Pos(';',st_enreg) - 1));
        st_trav2 := Trim(Copy(st_enreg, Pos(';',st_enreg)+1, 255));
        StringGridFichier.Cells[1, i_ind1] := st_trav1;
        st_enreg := st_trav2;
        StringGridFichier.Cells[2, i_ind1] := st_enreg;
        readln(Fichier, st_enreg);
        i_ind1 := Succ(i_ind1);
    end;
//  Ajout des parametres saisis à la liste du fichier .seq
    ParamGenCle := TStringList.Create;
    ParamGenData := TStringList.Create;
    st_Chemin := WorkingPath.Text;
    FicIni := TIniFile.Create(st_chemin + 'ListeFic');
    stDossier := FicIni.ReadString('DossierActif', 'Actif', '');
    if stDossier <> '' then stDossier := '_' + stDossier;
    FicIni.ReadSection('General' + stDossier, ParamGenCle);
    for i_ind2 := 0 to ParamGenCle.Count - 1 do
        begin
        StringGridFichier.Cells[0, i_ind1] := st_racine + 'PARAM.' + ParamGenCle.Strings[i_ind2];
        StringGridFichier.Cells[1, i_ind1] := st_racine + 'PARAM.' + ParamGenCle.Strings[i_ind2];
        StringGridFichier.Cells[2, i_ind1] := FicIni.ReadString('General' + stDossier, ParamGenCle.Strings[i_ind2], '');
        i_ind1 := Succ(i_ind1);
        end;
    FicIni.Free;

    if FileExists(st_Chemin + Edit1.Text + '.ini') then
    begin
        FicIni := TIniFile.Create(st_Chemin + Edit1.Text + '.ini');

        ListeAsso.Items.Text := '';
        FicIni.ReadSectionValues ('Fichier', ListeAsso.Items);
        i_ind1 := 0;
        while i_ind1 <= ListeAsso.Items.Count - 1 do
        begin
            st_trav1 := ListeAsso.Items.Strings[i_ind1];
            ReadTokenSt(st_trav1);
            StringGrid2.Cells[i_ind1 + 1, StringGrid1.Row] := ReadTokenSt(st_trav1);
            if StringGrid2.Cells[i_ind1 + 1, StringGrid1.Row] = '' then
                StringGrid2.Cells[i_ind1 + 1, StringGrid1.Row] := 'M';
            if i_ind1= 0 then
                StringGrid2.Cells[i_ind1 + 1, StringGrid1.Row] := 'C'
//  Modif 07/02/02 rapprochement GC - GRC
            else        
            StringGrid3.Cells[i_ind1 + 1, StringGrid1.Row] :=  Copy(ReadTokenSt(st_trav1), 0, 3);
//  Fin Modif 07/02/02 rapprochement GC - GRC
            i_ind1 := i_ind1 + 1;
        end;
        ListeAsso.Items.Text := '';
        for i_ind1 := 0 to StringGridAsso.RowCount do
        begin
            if StringGridAsso.cells[0, i_ind1] = '' then break;
            StringGridAsso.Rows[i_ind1].Clear;
            StringGridAssoStruct.Rows[i_ind1].Clear;
            StringGridAssoFic.Rows[i_ind1].Clear;
        end;
        i_ind1 := TabControl1.TabIndex;
//  recherche du fichier precedent en Creation
//  Modif 07/02/02 rapprochement GC - GRC
//        while (i_ind1 <> 0) and (StringGrid2.Cells[i_ind1 + 1, StringGrid1.Row] <> 'C') do
//            Dec (i_ind1);
//  chargement des sections jusqu'au suivant en creation ou a la fin
        i_ligne_stringgrid := 0;

        Charge_Section(FicIni, st_racine, i_ind1 + 1, i_ligne_stringgrid);
{        repeat
            st_racine := StringGrid3.Cells[i_ind1 + 1, StringGrid1.Row];
            Charge_Section(FicIni, st_racine, i_ind1 + 1, i_ligne_stringgrid);
            Inc (i_ind1);
        until (i_ind1 > TabControl1.Tabs.Count) or
              (StringGrid2.Cells[i_ind1 + 1, StringGrid1.Row] = 'C') or
              (StringGrid2.Cells[i_ind1 + 1, StringGrid1.Row] = ''); }
//  Fin Modif 07/02/02 rapprochement GC - GRC
        FicIni.Free;
    end;
//  Chargement de la liste des champs de la table dans ListeTable
    Tob1 := TOB.Create('Liste des champs',Nil,-1);
    Tob2 := TOB.Create('DECHAMPS',Nil,-1);
    Query1 := OpenSql('SELECT * FROM DETABLES WHERE DT_NOMTABLE = "' + Edit1.Text + '"', False);
    Tob1.LoadDetailDB('DETABLES','','',Query1,FALSE);
    Tob2 := Tob1.Detail[0] ;
    Tob2.LoadDetailDB('DECHAMPS','"'+Tob2.GetValue('DT_PREFIXE')+'"','DH_NUMCHAMP',Nil,FALSE) ;
    Tob2.PutStringsDetail(StringGridTable.Cols[0],'DH_NOMCHAMP');
    Tob2.PutStringsDetail(StringGridTable.Cols[1],'DH_LIBELLE');
    Ferme(Query1);
    Tob1.Free;

    ListeDetailAsso;
    ListeDictFichier;
    RadioGroup1.ItemIndex := 0;
    RadioGroup1Click(Self);
    i_AssoRow := ListeAsso.Items.Count - 1;
    i_AssoCol := 0;
end;

//----------------------------------------------------------------------------------
function TFAssocImport.VerifCorrespondance (i_nb_champs : integer;
                                            st_champs, st_racine : string;
                                            var i_nrow : integer) : boolean;
var
    i_ind : integer;
    b_trouve : boolean;
begin
    b_trouve := False;
    i_nrow := 0;
    if (st_champs[1] = '"') or (Copy(st_champs, 0, 3) <> st_racine) then
    begin
        b_trouve := True;
        i_nrow := -1;
    end
    else
    begin
        for i_ind := 0 to i_nb_champs do
        begin
            if (StringGridFichier.Cells[0, i_ind] = st_champs) then
            begin
                b_trouve := True;
                i_nrow := i_ind;
                break;
            end;
        end;
    end;
    Result := b_trouve;
end;

//----------------------------------------------------------------------------------
procedure TFAssocImport.Charge_Section(FicIni : TIniFile; st_racine : string;
                                       i_col : integer; var i_ligne : integer);
var
    st_enreg, st_trav1, st_trav2, st_file, st_maj : string;
    i_ind1, i_ind2, i_ind3, i_ind4, i_indtext, i_nrow : integer;
    i_nb_ch_text, i_ligne_stringgrid : integer;
    b_Trouve : boolean;
    Fichier : TextFile;
    FicParam : TIniFile;
    ParamGenCle : TStrings;
    ParamGenData : TStrings;

begin
    for i_ind1 := 0 to StringGridFichier.RowCount - 1 do
    begin
        if StringGridFichier.cells[0, i_ind1] = '' then break;
        StringGridFichier.Rows[i_ind1].Clear;
    end;
//  Ouverture dictionnaire fichier texte
    AssignFile(Fichier, StringGrid1.Cells[i_col, StringGrid1.Row] + '.des');
    Reset(Fichier);
//  Chargement de la liste des champs du fichier texte
    readln(Fichier, st_enreg);
    i_ind1 := 0;
    i_ind3 := 0;
    while st_enreg <> '' do
    begin
        i_ind2 := 0;
        st_trav1 := Trim(Copy(st_enreg, 1, Pos(';',st_enreg) - 1));
        st_trav2 := Trim(Copy(st_enreg, Pos(';',st_enreg)+1, 255));
        StringGridFichier.Cells[0, i_ind1] := st_trav1;
        st_enreg := st_trav2;
        st_trav1 := Trim(Copy(st_enreg, 1, Pos(';',st_enreg) - 1));
        st_trav2 := Trim(Copy(st_enreg, Pos(';',st_enreg)+1, 255));
        StringGridFichier.Cells[1, i_ind1] := st_trav1;
        st_enreg := st_trav2;
        StringGridFichier.Cells[2, i_ind1] := st_enreg;
        readln(Fichier, st_enreg);
        i_ind1 := Succ(i_ind1);
    end;
//  Ajout des parametres saisis à la liste du fichier .seq
    ParamGenCle := TStringList.Create;
    ParamGenData := TStringList.Create;
    st_Chemin := WorkingPath.Text;
    FicParam := TIniFile.Create(st_chemin + 'ListeFic');
    FicParam.ReadSection('General' + stDossier, ParamGenCle);
    for i_ind2 := 0 to ParamGenCle.Count - 1 do
        begin
        StringGridFichier.Cells[0, i_ind1] := copy(st_racine,0,3) + 'PARAM.' + ParamGenCle.Strings[i_ind2];
        StringGridFichier.Cells[1, i_ind1] := copy(st_racine,0,3) + 'PARAM.' + ParamGenCle.Strings[i_ind2];
        StringGridFichier.Cells[2, i_ind1] := FicParam.ReadString('General' + stDossier, ParamGenCle.Strings[i_ind2], '');
        i_ind1 := Succ(i_ind1);
        end;
    FicParam.Free;

    i_nb_ch_text := i_ind1;

    ListeAsso.Items.Text := '';
    ListeAssoCh.Items.Text := '';
    ListeAssoStruct.Items.Text := '';

//  Modif 07/02/02 rapprochement GC - GRC
    St_maj := '';
    if ((st_racine = 'PRO') or ((Edit1.text = 'TIERSCOMPL') and (st_racine = 'GRI')))
    then St_maj := StringGrid2.Cells[TabControl1.TabIndex + 1, StringGrid1.Row];

    if (StringGrid2.Cells[TabControl1.TabIndex + 1, StringGrid1.Row] = 'M') and
       ((Pos('=', Edit4.Text) <> 0) or (Edit4.Text = '')) then
        Edit4.Text := FicIni.ReadString( 'Cle_' + st_racine+ st_maj,'Cle','');

    FicIni.ReadSectionValues('Champs_' + st_racine + st_maj,ListeAssoCh.Items);
    FicIni.ReadSectionValues('Designation_' + st_racine + st_maj,ListeAsso.Items);
    FicIni.ReadSectionValues('Offset_' + st_racine + st_maj,ListeAssoStruct.Items);
//  Modif Fin 07/02/02 rapprochement GC - GRC
    i_ind1 := 0;
    i_indtext := 0;
    while i_indtext <= ListeAssoCh.Items.Count - 1 do
    begin
        b_trouve := True; // valeur initiale
             // mise à jour champs et Offset
        st_trav1 := ListeAssoCh.Items.Strings[i_indtext];
        st_trav2 := ListeAssoStruct.Items.Strings[i_indtext];
        StringGridAssoFic.Cells[0, i_ligne] := Copy(st_trav1, 0, Pos('=',st_trav1) - 1);
        StringGridAssoStruct.Cells[0, i_ligne] := Copy(st_trav2, 0, Pos('=',st_trav2) - 1);
        st_trav1 := Copy(st_trav1, Pos('=',st_trav1) + 1, 255);
        st_trav2 := Copy(st_trav2, Pos('(',st_trav2) + 1, 255);
        if Pos('{SI', st_trav1) <> 0 then
            begin
            st_trav1 := ListeAssoCh.Items.Strings[i_indtext];
            StringGridAsso.Cells[0, i_ligne] := Copy(st_trav1, 0, Pos('=',st_trav1) - 1);
            st_trav1 := Copy(st_trav1, Pos('=',st_trav1) + 1, 255);
            StringGridAsso.Cells[1, i_ligne] := st_trav1;
            StringGridAssoFic.Cells[1, i_ligne] := st_trav1;
            StringGridAssoStruct.Cells[1, i_ligne] := st_trav1;
            i_ligne := i_ligne + 1;
            end
            else
            begin
            i_ind2 := 1;
            while Pos('+', st_trav1) <> 0 do
            begin
                    //Verif code champs avec fichier texte
                b_trouve := VerifCorrespondance (i_nb_ch_text,
                                    Copy(st_trav1, 0, Pos('+',st_trav1) - 1), st_racine, i_nrow);
                if b_trouve = true then
                begin
                    StringGridAssoFic.Cells[i_ind2, i_ligne] := Copy(st_trav1, 0, Pos('+',st_trav1) - 1);
                    st_trav1 := Copy(st_trav1, Pos('+',st_trav1) + 1, 255);
                    if i_nrow = -1 then
                    begin
                        StringGridAssoStruct.Cells[i_ind2, i_ligne] := Copy(st_trav2, 0, Pos(')',st_trav2) - 1);
                    end
                    else
                    begin
                        StringGridAssoStruct.Cells[i_ind2, i_ligne] := StringGridFichier.Cells[2, i_nrow];
                    end;
                    st_trav2 := Copy(st_trav2, Pos('(',st_trav2) + 1, 255);
                    i_ind2 := i_ind2 + 1;
                end
                else
                begin
                    break;
                end;
            end;
            if b_trouve = true then
            begin
                    // Verif code champs avec fichier texte
                b_trouve := VerifCorrespondance (i_nb_ch_text, st_trav1, st_racine, i_nrow);
                if b_trouve = False then
                begin
                    ListeAsso.Items.Delete(i_indtext);
                    ListeAssoCh.Items.Delete(i_indtext);
                    ListeAssoStruct.Items.Delete(i_indtext);
                    i_indtext := i_indtext + 1;
                    continue;
                end;
                StringGridAssoFic.Cells[i_ind2, i_ligne] := st_trav1;
                if i_nrow = -1 then
                begin
                    StringGridAssoStruct.Cells[i_ind2, i_ligne] := Copy(st_trav2, 0, Pos(')',st_trav2) - 1);
                end
                else
                begin
                    StringGridAssoStruct.Cells[i_ind2, i_ligne] := StringGridFichier.Cells[2, i_nrow];
                end;

                st_trav1 := ListeAsso.Items.Strings[i_indtext];
                StringGridAsso.Cells[0, i_ligne] := Copy(st_trav1, 0, Pos('=',st_trav1) - 1);
                st_trav1 := Copy(st_trav1, Pos('=',st_trav1) + 1, 255);
                i_ind2 := 1;
                while Pos('+', st_trav1) <> 0 do
                begin
                    StringGridAsso.Cells[i_ind2, i_ligne] := Copy(st_trav1, 0, Pos('+',st_trav1) - 1);
                    st_trav1 := Copy(st_trav1, Pos('+',st_trav1) + 1, 255);
                    i_ind2 := i_ind2 + 1;
                end;
                StringGridAsso.Cells[i_ind2, i_ligne] := st_trav1;

                i_ligne := i_ligne + 1;
            end;
            end;
        i_indtext := i_indtext + 1;
    end;
    ListeAsso.Items.Text := '';
    ListeAssoCh.Items.Text := '';
    ListeAssoStruct.Items.Text := '';
end;

//----------------------------------------------------------------------------------
procedure TFAssocImport.BitBtn1Click(Sender: TObject);
var
    i_ind1 : integer;

begin
    i_AssoRow := i_AssoRow + 1;
    i_AssoCol := 0;
    for i_ind1 := 0 to StringGridTable.RowCount do
        if StringGridTable.Cells[RadioGroup1.ItemIndex, i_ind1] = ListeTable.Items.Strings[ListeTable.ItemIndex] then
            Break;
    StringGridAsso.Cells[i_AssoCol, i_AssoRow] := StringGridTable.Cells[0, i_ind1];
    StringGridAssoStruct.Cells[i_AssoCol, i_AssoRow] := StringGridTable.Cells[0, i_ind1];
    StringGridAssoFic.Cells[i_AssoCol, i_AssoRow] := StringGridTable.Cells[0, i_ind1];
    i_AssoCol := i_AssoCol + 1;
    ListeDetailAsso;
    RadioGroup1Click(Self);
    ListeAsso.ItemIndex := ListeAsso.Items.Count - 1;
    if ListeAsso.ItemIndex <> -1 then
        case RadioGroup2.ItemIndex of
        0 : Edit3.Text := ListeAsso.Items.Strings[ListeAsso.ItemIndex];
        1 : Edit3.Text := ListeAssoStruct.Items.Strings[ListeAssoStruct.ItemIndex];
        end;
end;

//----------------------------------------------------------------------------------
procedure TFAssocImport.BitBtn3Click(Sender: TObject);
var
    i_ind1, i_ind2, i_ind3, i_ind4, i_ItemIndexSave : integer;
    s_trav1 : string;

begin
//  Pas de ligne selectionnee au centre
    if ListeAsso.ItemIndex < 0 then Exit;
//  Recupere le nom du champ dans la ligne centrale
    i_ItemIndexSave := ListeAsso.ItemIndex;
    if Pos('=',ListeAsso.Items.Strings[ListeAsso.ItemIndex]) <> 0 then
        s_trav1 := Copy(ListeAsso.Items.Strings[ListeAsso.ItemIndex], 0,
                        Pos('=',ListeAsso.Items.Strings[ListeAsso.ItemIndex]) - 1)
    else
        s_trav1 := ListeAsso.Items.Strings[ListeAsso.ItemIndex];
//  Recherche de la ligne selection dans la StringGrid
//    i_ind1 := -1;
//    repeat
//        i_ind1 := i_ind1 + 1;
//    until (StringGridAsso.Cells[0, i_ind1] = s_trav1) or
//          (i_ind1 > StringGridAsso.RowCount);
//    if i_ind1 > StringGridAsso.RowCount then Exit;
    i_ind1 := ListeAsso.ItemIndex;
    if i_ind1 <= StringGridAsso.RowCount then
    begin
//  Recherche de la premiere colonne libre dans la ligne
        i_ind2 := -1;
        repeat
            i_ind2 := i_ind2 + 1;
        until (StringGridAsso.Cells[i_ind2, i_ind1] = '') or
              (i_ind2 > StringGridAsso.ColCount);
        if i_ind2 > StringGridAsso.ColCount then Exit;
        if i_ind2 <= StringGridAsso.ColCount then
            for i_ind3 := 0 to ListeFichier.Items.Count - 1 do
                if ListeFichier.Selected[i_ind3] then
                begin
                    StringGridAsso.Cells[i_ind2, i_ind1] := ListeFichier.Items.Strings[i_ind3];
                    StringGridAssoStruct.Cells[i_ind2, i_ind1] := StringGridFichier.Cells[2, i_ind3];
                    StringGridAssoFic.Cells[i_ind2, i_ind1] := StringGridFichier.Cells[0, i_ind3];
                end;
    end;
    ListeDetailAsso;
    ListeAsso.ItemIndex := i_ItemIndexSave;
//    ListeDictFichier;
    if ListeAsso.ItemIndex <> -1 then
        case RadioGroup2.ItemIndex of
        0 : Edit3.Text := ListeAsso.Items.Strings[ListeAsso.ItemIndex];
        1 : Edit3.Text := ListeAssoStruct.Items.Strings[ListeAssoStruct.ItemIndex];
        end;
end;

//----------------------------------------------------------------------------------
procedure TFAssocImport.BitBtn2Click(Sender: TObject);
var
    i_ind1, i_ind2, i_ind3 : integer;

begin
    i_ind1 := ListeAsso.ItemIndex;
    i_ind2 := i_ind1 + 1;
    while StringGridAsso.Cells[0, i_ind1] <> '' do
    begin
        i_ind3 := 0;
        while i_ind3 <= StringGridAsso.ColCount do
        begin
            StringGridAsso.Cells[i_ind3, i_ind1] := StringGridAsso.Cells[i_ind3, i_ind2];
            StringGridAssoStruct.Cells[i_ind3, i_ind1] := StringGridAssoStruct.Cells[i_ind3, i_ind2];
            StringGridAssoFic.Cells[i_ind3, i_ind1] := StringGridAssoFic.Cells[i_ind3, i_ind2];
            if i_ind3 <> 0 then
            begin
                StringGridAsso.Cells[i_ind3, i_ind2] := '';
                StringGridAssoStruct.Cells[i_ind3, i_ind2] := '';
                StringGridAssoFic.Cells[i_ind3, i_ind2] := '';
            end;
            i_ind3 := i_ind3 + 1;
        end;
        i_ind1 := i_ind1 + 1;
        i_ind2 := i_ind2 + 1;
    end;
    i_AssoRow := i_AssoRow - 1;
    StringGridAsso.Cells[0, i_ind1 - 1] := '';
    StringGridAssoStruct.Cells[0, i_ind1 - 1] := '';
    StringGridAssoFic.Cells[0, i_ind1 - 1] := '';
    ListeAsso.Items.Delete(ListeAsso.ItemIndex);
    ListeAssoStruct.Items.Delete(ListeAsso.ItemIndex);
    ListeDictFichier;
    RadioGroup1Click(Self);
    ListeDetailAsso;
    if ListeAsso.ItemIndex <> -1 then
        case RadioGroup2.ItemIndex of
        0 : Edit3.Text := ListeAsso.Items.Strings[ListeAsso.ItemIndex];
        1 : Edit3.Text := ListeAssoStruct.Items.Strings[ListeAssoStruct.ItemIndex];
        end;
end;

//----------------------------------------------------------------------------------
procedure TFAssocImport.BitBtn4Click(Sender: TObject);
var
    i_ind1, i_ind2, i_ind3 : integer;

begin
    i_ind1 := ListeAsso.ItemIndex;
//  Modif 06/02/02 rapprochement GC - GRC
    if i_ind1 < 0 then Exit;
//  Fin Modif 06/02/02 rapprochement GC - GRC
    i_ind2 := 1;
    while i_ind2 <= StringGridAsso.ColCount do
    begin
        StringGridAsso.Cells[i_ind2, i_ind1] := '';
        StringGridAssoStruct.Cells[i_ind2, i_ind1] := '';
        StringGridAssoFic.Cells[i_ind2, i_ind1] := '';
        i_ind2 := i_ind2 + 1;
    end;
    ListeDictFichier;
    RadioGroup1Click(Self);
    ListeDetailAsso;
    if ListeAsso.ItemIndex <> -1 then
        case RadioGroup2.ItemIndex of
        0 : Edit3.Text := ListeAsso.Items.Strings[ListeAsso.ItemIndex];
        1 : Edit3.Text := ListeAssoStruct.Items.Strings[ListeAssoStruct.ItemIndex];
        end;
end;

//----------------------------------------------------------------------------------
procedure TFAssocImport.BitBtn5Click(Sender: TObject);
var
    i_ind1, i_ind2, i_ItemIndexSave : integer;
    s_trav1 : string;

begin
//  Pas de ligne selectionnee au centre
    if ListeAsso.ItemIndex < 0 then Exit;
    FSaisieTexte.Left := (Left + Trunc(Width / 2)) - Trunc(FSaisieTexte.Width / 2);
    FSaisieTexte.Top := (Top + Trunc(Height / 2)) - Trunc(FSaisieTexte.Height / 2);
    FSaisieTexte.Edit1.Text := '';
    FSaisieTexte.ShowModal;
    if FSaisieTexte.Edit1.Text <> '' then
    begin
//  Recupere le nom du champ dans la ligne centrale
        i_ind1 := ListeAsso.ItemIndex;
{        if Pos('=',ListeAsso.Items.Strings[ListeAsso.ItemIndex]) <> 0 then
            s_trav1 := Copy(ListeAsso.Items.Strings[ListeAsso.ItemIndex], 0,
                            Pos('=',ListeAsso.Items.Strings[ListeAsso.ItemIndex]) - 1)
        else
            s_trav1 := ListeAsso.Items.Strings[ListeAsso.ItemIndex];  }
//  Recherche de la ligne selection dans la StringGrid
{        i_ind1 := -1;
        repeat
            i_ind1 := i_ind1 + 1;
        until (StringGridAsso.Cells[0, i_ind1] = s_trav1) or
              (i_ind1 > StringGridAsso.RowCount);
        if i_ind1 > StringGridAsso.RowCount then Exit;  }
        if i_ind1 <= StringGridAsso.RowCount then
        begin
//  Recherche de la premiere colonne libre dans la ligne
            i_ind2 := -1;
            repeat
                i_ind2 := i_ind2 + 1;
            until (StringGridAsso.Cells[i_ind2, i_ind1] = '') or
                  (i_ind2 > StringGridAsso.ColCount);
            if i_ind2 > StringGridAsso.ColCount then Exit;
            if i_ind2 <= StringGridAsso.ColCount then
            begin
                StringGridAsso.Cells[i_ind2, i_ind1] := '"' + FSaisieTexte.Edit1.Text + '"';
                StringGridAssoStruct.Cells[i_ind2, i_ind1] := '"' + FSaisieTexte.Edit1.Text + '"';
                StringGridAssoFic.Cells[i_ind2, i_ind1] := '"' + FSaisieTexte.Edit1.Text + '"';
            end;
        end;
        ListeDetailAsso;
        ListeAsso.ItemIndex := i_ItemIndexSave;
        ListeDictFichier;
        if ListeAsso.ItemIndex <> -1 then
            case RadioGroup2.ItemIndex of
            0 : Edit3.Text := ListeAsso.Items.Strings[ListeAsso.ItemIndex];
            1 : Edit3.Text := ListeAssoStruct.Items.Strings[ListeAssoStruct.ItemIndex];
            end;
    end;
end;

//----------------------------------------------------------------------------------
procedure TFAssocImport.BitBtn6Click(Sender: TObject);
var
    i_ind1, i_ind2 : integer;

begin
    i_AssoRow := i_AssoRow + 1;
    i_AssoCol := 0;
//    for i_ind1 := 0 to StringGridTable.RowCount do
//        if StringGridTable.Cells[RadioGroup1.ItemIndex, i_ind1] = ListeTable.Items.Strings[ListeTable.ItemIndex] then
//            Break;
    i_ind2 := 0;
    for i_ind1 := 0 to StringGridAsso.RowCount - 1 do
        if Copy(StringGridAsso.Cells[0, i_ind1], 0, 8) = 'Tablette' then
            Inc(i_ind2);
    StringGridAsso.Cells[i_AssoCol, i_AssoRow] := 'Tablette' + IntToStr(i_ind2 + 1);
    StringGridAssoStruct.Cells[i_AssoCol, i_AssoRow] := 'Tablette' + IntToStr(i_ind2 + 1);
    StringGridAssoFic.Cells[i_AssoCol, i_AssoRow] := 'Tablette' + IntToStr(i_ind2 + 1);
    i_AssoCol := i_AssoCol + 1;
    ListeDetailAsso;
    RadioGroup1Click(Self);
    ListeAsso.ItemIndex := ListeAsso.Items.Count - 1;
    if ListeAsso.ItemIndex <> -1 then
        case RadioGroup2.ItemIndex of
        0 : Edit3.Text := ListeAsso.Items.Strings[ListeAsso.ItemIndex];
        1 : Edit3.Text := ListeAssoStruct.Items.Strings[ListeAssoStruct.ItemIndex];
        end;
end;

//----------------------------------------------------------------------------------
procedure TFAssocImport.ListeTableDblClick(Sender: TObject);
begin
    BitBtn1Click(Sender);
end;

//----------------------------------------------------------------------------------
procedure TFAssocImport.ListeDictFichier;
var
    i_ind1 : integer;
    st_enreg : string;

begin
    ListeFichier.Items.Text := '';
    i_ind1 := 0;
    while StringGridFichier.Cells[RadioGroup4.ItemIndex, i_ind1] <> '' do
    begin
//        if Pos(StringGridFichier.Cells[1, i_ind1], ListeFichier.Items.Text) = 0 then
            if i_ind1 = 0 then
                ListeFichier.Items.Text := StringGridFichier.Cells[RadioGroup4.ItemIndex, i_ind1]
            else
                ListeFichier.Items.Add(StringGridFichier.Cells[RadioGroup4.ItemIndex, i_ind1]);
        i_ind1 := Succ(i_ind1);
    end;
end;

//----------------------------------------------------------------------------------
procedure TFAssocImport.ListeDetailAsso;
var
    i_ind1, i_ind2 : integer;

begin
    ListeAsso.Items.Text := '';
    ListeAssoStruct.Items.Text := '';
    ListeAssoCh.Items.Text := '';
    i_ind1 := 0;
    i_ind2 := 0;
    while StringGridAsso.Cells[0, i_ind1] <> '' do
    begin
        if i_ind2 = 0 then
        begin
            ListeAsso.Items.Add('');
            ListeAssoStruct.Items.Add('');
            ListeAssoCh.Items.Add('');
        end;
        ListeAsso.Items.Strings[i_ind1] := StringGridAsso.Cells[0, i_ind1];
        ListeAssoStruct.Items.Strings[i_ind1] := StringGridAsso.Cells[0, i_ind1];
        ListeAssoCh.Items.Strings[i_ind1] := StringGridAsso.Cells[0, i_ind1];
        i_ind2 := 1;
        while StringGridAsso.Cells[i_ind2, i_ind1] <> '' do
        begin
            if i_ind2 = 1 then
            begin
                ListeAsso.Items.Strings[i_ind1] := ListeAsso.Items.Strings[i_ind1] + '=' +
                                                   StringGridAsso.Cells[i_ind2, i_ind1];
                ListeAssoStruct.Items.Strings[i_ind1] := ListeAssoStruct.Items.Strings[i_ind1] + '=' +
                                                         '(' + StringGridAssoStruct.Cells[i_ind2, i_ind1] + ')';
                ListeAssoCh.Items.Strings[i_ind1] := ListeAssoCh.Items.Strings[i_ind1] + '=' +
                                                   StringGridAssoFic.Cells[i_ind2, i_ind1];
            end
            else
            begin
                ListeAsso.Items.Strings[i_ind1] := ListeAsso.Items.Strings[i_ind1] + '+' +
                                                   StringGridAsso.Cells[i_ind2, i_ind1];
                ListeAssoStruct.Items.Strings[i_ind1] := ListeAssoStruct.Items.Strings[i_ind1] + '+' +
                                                         '(' + StringGridAssoStruct.Cells[i_ind2, i_ind1] + ')';
                ListeAssoCh.Items.Strings[i_ind1] := ListeAssoCh.Items.Strings[i_ind1] + '+' +
                                                   StringGridAssoFic.Cells[i_ind2, i_ind1];
            end;
            i_ind2 := i_ind2 + 1;
        end;
        i_ind2 := 0;
        i_ind1 := Succ(i_ind1);
    end;
end;

//----------------------------------------------------------------------------------
procedure TFAssocImport.RadioGroup1Click(Sender: TObject);
var
    i_ind1, i_ind2 : integer;
    bTrouve : boolean;

begin
    ListeTable.Items.Text := '';
    i_ind1 := 0;
    while StringGridTable.Cells[RadioGroup1.ItemIndex, i_ind1] <> '' do
    begin
        i_ind2 := 0;
        bTrouve := False;
        if ListeAsso.Items.Text <> '' then
            while i_ind2 <= ListeAsso.Items.Count - 1 do
            begin
                if StringGridTable.Cells[0, i_ind1] = ListeAsso.Items.Strings[i_ind2] then
                    bTrouve := True;
                i_ind2 := i_ind2 + 1;
            end;
//        if Pos(StringGridTable.Cells[0, i_ind1], ListeAsso.Items.Text) = 0 then
        if not (bTrouve) then
            if i_ind1 = 0 then
                ListeTable.Items.Text := StringGridTable.Cells[RadioGroup1.ItemIndex, i_ind1]
            else
                ListeTable.Items.Add(StringGridTable.Cells[RadioGroup1.ItemIndex, i_ind1]);
        i_ind1 := Succ(i_ind1);
    end;
end;

//----------------------------------------------------------------------------------
procedure TFAssocImport.ListeFichierDblClick(Sender: TObject);
begin
    BitBtn3Click(Sender);
end;

//----------------------------------------------------------------------------------
procedure TFAssocImport.StringGrid1DblClick(Sender: TObject);
var
    i_ind1 : integer;
    st_trav1, st_file : string;
    Fichier: textfile;

begin
    Enregistre_Inifile(i_TabIndexPrec + 1, i_RowPrec);
    i_TabIndexPrec := 0;
    i_RowPrec := StringGrid1.Row;
    Edit1.Text := StringGrid1.Cells[0, StringGrid1.Row];
    Edit2.Text := StringGrid1.Cells[1, StringGrid1.Row];
    st_Chemin := WorkingPath.Text;
    st_file := st_Chemin + Edit1.Text;
    TabControl1.Tabs.Clear;
    for i_ind1 := 1 to 5 do
    begin
        if StringGrid1.Cells[i_ind1, StringGrid1.Row] <> '' then
        begin
            st_trav1 := ExtractFileName(StringGrid1.Cells[i_ind1, StringGrid1.Row]);
            TabControl1.Tabs.Append(st_trav1);
            if not FileExists (st_file + '.ini') then
            begin
                AssignFile(Fichier, StringGrid1.Cells[i_ind1, StringGrid1.Row] + '.des');
                Reset(Fichier);
                readln(Fichier, st_trav1);
                st_trav1 := Copy(st_trav1, 0, 3);
                CloseFile(Fichier);
                if i_ind1 > 1 then
                    StringGrid2.Cells[i_ind1, StringGrid1.Row] := 'M'
                else
                    StringGrid2.Cells[i_ind1, StringGrid1.Row] := 'C';
                StringGrid3.Cells[i_ind1, StringGrid1.Row] := st_trav1;
            end;
        end;
    end;
    ListeAsso.Items.Text := '';
    ListeAssoStruct.Items.Text := '';
    ListeAssoCh.Items.Text := '';
    ListeFichier.Items.Text := '';
    ListeTable.Items.Text := '';
    for i_ind1 := 0 to StringGridAsso.RowCount do
    begin
        if StringGridAsso.cells[0, i_ind1] = '' then break;
        StringGridAsso.Rows[i_ind1].Clear;
        StringGridAssoStruct.Rows[i_ind1].Clear;
        StringGridAssoFic.Rows[i_ind1].Clear;
        StringGridFichier.Rows[i_ind1].Clear;
        StringGridTable.Rows[i_ind1].Clear;
    end;
    Edit1.Text := StringGrid1.Cells[0, StringGrid1.Row];
    Edit2.Text := StringGrid1.Cells[1, StringGrid1.Row];
    if Edit2.Text <> '' then
    begin
        RadioGroup3.ItemIndex := 0;
        RadioGroup3.Enabled := false;
        Charge_Dictionnaire;
    end
end;

//----------------------------------------------------------------------------------
procedure TFAssocImport.RadioGroup2Click(Sender: TObject);
begin
    case RadioGroup2.ItemIndex of
    0 :
        begin
            ListeAsso.Visible := True;
            ListeAssoCh.Visible := False;
            ListeAssoStruct.Visible := False;
            if ListeAsso.ItemIndex <> -1 then
                Edit3.Text := ListeAsso.Items.Strings[ListeAsso.ItemIndex];
        end;
    1 :
        begin
            ListeAsso.Visible := False;
            ListeAssoCh.Visible := False;
            ListeAssoStruct.Visible := True;
            if ListeAssoStruct.ItemIndex <> -1 then
                Edit3.Text := ListeAssoStruct.Items.Strings[ListeAssoStruct.ItemIndex];
        end;
    2 :
        begin
            ListeAsso.Visible := False;
            ListeAssoCh.Visible := True;
            ListeAssoStruct.Visible := False;
            if ListeAssoCh.ItemIndex <> -1 then
                Edit3.Text := ListeAssoCh.Items.Strings[ListeAssoCh.ItemIndex];
        end;
    end;
end;

//----------------------------------------------------------------------------------
procedure TFAssocImport.Enregistre_Inifile(i_col,i_row : integer);
var
    i_ind1, i_ind2, i_ind3 : integer;
    st_trav1, st_trav2, st_racine : string;
    FicIni: TIniFile;

begin
    if Edit2.Text <> '' then
    begin
        st_Chemin := WorkingPath.Text;
//        if not MAJ then DeleteFile(st_Chemin + Edit1.Text + '.ini');
        FicIni := TIniFile.Create(st_Chemin + Edit1.Text + '.ini');
//  Modif 06/02/02 rapprochement GC - GRC
        st_racine := StringGrid3.Cells [i_col,i_row];
        if ((st_racine = 'PRO') or ((Edit1.text = 'TIERSCOMPL') and (st_racine = 'GRI'))) then
            st_racine := st_racine + StringGrid2.Cells [i_col,i_row];
//  Fin Modif 06/02/02 rapprochement GC - GRC
        FicIni.EraseSection('Cle_' + st_racine);
        FicIni.EraseSection('Champs_' + st_racine);
        FicIni.EraseSection('Designation_' + st_racine);
        FicIni.EraseSection('Offset_' + st_racine);

        FicIni.WriteString('Cle_' + st_racine, 'Cle', Edit4.Text);

        for i_ind1 := 0 to ListeAsso.Items.Count - 1 do
        begin
            st_trav1 := Copy(ListeAssoCh.Items.Strings[i_ind1],
                             0, Pos('=',ListeAssoCh.Items.Strings[i_ind1]) - 1);
            st_trav2 := Copy(ListeAssoCh.Items.Strings[i_ind1],
                             Pos('=',ListeAssoCh.Items.Strings[i_ind1]) + 1, 255);
            if (Pos('"', st_trav2) = 1) or (Pos('"', st_trav2) = Length(st_trav2)) then
                st_trav2 := '"' + st_trav2 + '"'
            else
                if Pos('{SI', ListeAssoCh.Items.Strings[i_ind1]) = 0 then
                    st_racine := Copy(ListeAssoCh.Items.Strings[i_ind1],
                                      Pos('=',ListeAssoCh.Items.Strings[i_ind1]) + 1, 3);
//  Modif 06/02/02 rapprochement GC - GRC
            if ((st_racine = 'PRO') or ((Edit1.text = 'TIERSCOMPL') and (st_racine = 'GRI'))) then
                st_racine := st_racine + StringGrid2.Cells [i_col,i_row];
//  Fin Modif 06/02/02 rapprochement GC - GRC
            FicIni.WriteString('Champs_' + st_racine, st_trav1, st_trav2);
        end;
        for i_ind1 := 0 to ListeAsso.Items.Count - 1 do
        begin
            st_trav1 := Copy(ListeAsso.Items.Strings[i_ind1],
                             0, Pos('=',ListeAsso.Items.Strings[i_ind1]) - 1);
            st_trav2 := Copy(ListeAsso.Items.Strings[i_ind1],
                             Pos('=',ListeAsso.Items.Strings[i_ind1]) + 1, 255);
            if (Pos('"', st_trav2) = 1) or (Pos('"', st_trav2) = Length(st_trav2)) then
                st_trav2 := '"' + st_trav2 + '"'
            else
                if Pos('{SI', ListeAssoCh.Items.Strings[i_ind1]) = 0 then
                    st_racine := Copy(ListeAssoCh.Items.Strings[i_ind1],
                                      Pos('=',ListeAssoCh.Items.Strings[i_ind1]) + 1, 3);
//  Modif 06/02/02 rapprochement GC - GRC
            if ((st_racine = 'PRO') or ((Edit1.text = 'TIERSCOMPL') and (st_racine = 'GRI'))) then
                st_racine := st_racine + StringGrid2.Cells [i_col,i_row];
//  Fin Modif 06/02/02 rapprochement GC - GRC
            FicIni.WriteString('Designation_' + st_racine, st_trav1, st_trav2);
        end;
        for i_ind1 := 0 to ListeAsso.Items.Count - 1 do
        begin
            st_trav1 := Copy(ListeAssoStruct.Items.Strings[i_ind1],
                             0, Pos('=',ListeAssoStruct.Items.Strings[i_ind1]) - 1);
            st_trav2 := Copy(ListeAssoStruct.Items.Strings[i_ind1],
                             Pos('=',ListeAssoStruct.Items.Strings[i_ind1]) + 1, 255);
            if (Pos('"', st_trav2) <> 2) and (Pos('"', st_trav2) <> Length(st_trav2) - 1) then
                if Pos('{SI', ListeAssoCh.Items.Strings[i_ind1]) = 0 then
                    st_racine := Copy(ListeAssoCh.Items.Strings[i_ind1],
                                      Pos('=',ListeAssoCh.Items.Strings[i_ind1]) + 1, 3);
//  Modif 06/02/02 rapprochement GC - GRC
            if ((st_racine = 'PRO') or ((Edit1.text = 'TIERSCOMPL') and (st_racine = 'GRI'))) then
                st_racine := st_racine + StringGrid2.Cells [i_col,i_row];
//  Fin Modif 06/02/02 rapprochement GC - GRC
            FicIni.WriteString('Offset_' + st_racine, st_trav1, st_trav2);
        end;
        for i_ind1 := 0 to StringGrid1.RowCount do
        begin
            if StringGrid1.Cells[0, i_ind1] = Edit1.Text then
            begin
                for i_ind2 := 1 to StringGrid1.ColCount do
                begin
                    if StringGrid1.Cells[i_ind2, i_ind1] <> '' then
                    begin
                        st_trav1 := 'Fichier ' + IntToStr(i_ind2);
                        st_trav2 := StringGrid1.Cells[i_ind2, i_ind1];
                        st_trav2 := st_trav2 + ';' + StringGrid2.Cells[i_ind2, i_ind1];
                        st_trav2 := st_trav2 + ';' + StringGrid3.Cells[i_ind2, i_ind1];
                        FicIni.WriteString('Fichier', st_trav1, st_trav2);
                    end;
                end;
            end;
        end;
        FicIni.Free;
    end;
end;

//----------------------------------------------------------------------------------
procedure TFAssocImport.TabControl1Change(Sender: TObject);
var
    i_ind1 : Integer;

begin
    Enregistre_Inifile(i_TabIndexPrec + 1, StringGrid1.Row);
    i_ind1 := TabControl1.TabIndex;
    Edit1.Text := StringGrid1.Cells[0, StringGrid1.Row];
    Edit2.Text := StringGrid1.Cells[i_ind1 + 1, StringGrid1.Row];
    RadioGroup3.Enabled := True;
    if TabControl1.TabIndex = 0 then
    begin
        StringGrid2.Cells[i_ind1 + 1, StringGrid1.Row] := 'C';
        RadioGroup3.Enabled := False
    end else
    begin
      if ((StringGrid2.Cells[i_ind1 + 1, StringGrid1.Row] <> 'M') and
      (StringGrid2.Cells[i_ind1 + 1, StringGrid1.Row] <> 'C')) then
      StringGrid2.Cells[i_ind1 + 1, StringGrid1.Row] := 'M';
    end;

    if StringGrid2.Cells[i_ind1 + 1, StringGrid1.Row] = 'C' then
    begin
      Edit4.Text := '';
      RadioGroup3.ItemIndex := 0
    end else
      RadioGroup3.ItemIndex := 1;

    i_TabIndexPrec := TabControl1.TabIndex;
    Charge_Dictionnaire;
end;

procedure TFAssocImport.RadioGroup3Click(Sender: TObject);
var
    i_ind1 : integer;

begin
    case RadioGroup3.ItemIndex of
    0 : StringGrid2.Cells[TabControl1.TabIndex + 1, StringGrid1.Row] := 'C';
    1 : StringGrid2.Cells[TabControl1.TabIndex + 1, StringGrid1.Row] := 'M';
    end;
    if (TabControl1.TabIndex <> 0) and
       (StringGrid2.Cells[TabControl1.TabIndex + 1, StringGrid1.Row] = 'C') then
    begin
        Enregistre_Inifile(i_TabIndexPrec + 1, StringGrid1.Row);
        for i_ind1 := 0 to StringGridAsso.RowCount - 1 do
        begin
            if StringGridAsso.cells[0, i_ind1] = '' then break;
            StringGridAsso.Rows[i_ind1].Clear;
            StringGridAssoStruct.Rows[i_ind1].Clear;
            StringGridAssoFic.Rows[i_ind1].Clear;
        end;
        Charge_Dictionnaire;
    end;
end;

procedure TFAssocImport.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    Enregistre_Inifile(i_TabIndexPrec + 1, StringGrid1.Row);
end;

procedure TFAssocImport.FormCreate(Sender: TObject);
begin
    FSaisieTexte := TFSaisieTexte.Create(Application);
    FSaisieFormule := TFSaisieFormule.Create(Application);
    Edit3.Text := '';
end;

procedure TFAssocImport.ListeAssoClick(Sender: TObject);
begin
    Maj_Edit3;
end;

procedure TFAssocImport.ListeAssoChClick(Sender: TObject);
begin
    Maj_Edit3;
end;

procedure TFAssocImport.ListeAssoStructClick(Sender: TObject);
begin
    Maj_Edit3;
end;

procedure TFAssocImport.Maj_Edit3;
begin
    if ListeAsso.Visible then
    begin
        Edit3.Text := ListeAsso.Items.Strings[ListeAsso.ItemIndex];
        ListeAssoCh.ItemIndex := ListeAsso.ItemIndex;
        ListeAssoStruct.ItemIndex := ListeAsso.ItemIndex;
    end;
    if ListeAssoCh.Visible then
    begin
        Edit3.Text := ListeAssoCh.Items.Strings[ListeAssoCh.ItemIndex];
        ListeAsso.ItemIndex := ListeAssoCh.ItemIndex;
        ListeAssoStruct.ItemIndex := ListeAssoCh.ItemIndex;
    end;
    if ListeAssoStruct.Visible then
    begin
        Edit3.Text := ListeAssoStruct.Items.Strings[ListeAssoStruct.ItemIndex];
        ListeAsso.ItemIndex := ListeAssoStruct.ItemIndex;
        ListeAssoCh.ItemIndex := ListeAssoStruct.ItemIndex;
    end;

end;
procedure TFAssocImport.BitBtn7Click(Sender: TObject);
var
    i_ind1 : integer;

begin
    if ListeAsso.Visible then i_ind1 := ListeAsso.ItemIndex;
    if ListeAssoCh.Visible then i_ind1 := ListeAssoCh.ItemIndex;
    if ListeAssoStruct.Visible then i_ind1 := ListeAssoStruct.ItemIndex;
    ListeAsso.Items.Exchange(i_ind1, i_ind1 - 1);
    ListeAssoCh.Items.Exchange(i_ind1, i_ind1 - 1);
    ListeAssoStruct.Items.Exchange(i_ind1, i_ind1 - 1);
    StringGridAsso.Rows[0].Exchange(i_ind1, i_ind1 - 1);
    StringGridAssoFic.Rows[0].Exchange(i_ind1, i_ind1 - 1);
    StringGridAssoStruct.Rows[0].Exchange(i_ind1, i_ind1 - 1);
//    if ListeAsso.Visible then ListeAsso.ItemIndex := i_ind1 - 1;
//    if ListeAssoCh.Visible then ListeAssoCh.ItemIndex := i_ind1 - 1;
//    if ListeAssoStruct.Visible then ListeAssoStruct.ItemIndex := i_ind1 - 1;
    ListeAsso.ItemIndex := i_ind1 - 1;
    ListeAssoCh.ItemIndex := i_ind1 - 1;
    ListeAssoStruct.ItemIndex := i_ind1 - 1;
    Maj_Edit3;
end;

procedure TFAssocImport.BitBtn8Click(Sender: TObject);
var
    i_ind1 : integer;

begin
    if ListeAsso.Visible then i_ind1 := ListeAsso.ItemIndex;
    if ListeAssoCh.Visible then i_ind1 := ListeAssoCh.ItemIndex;
    if ListeAssoStruct.Visible then i_ind1 := ListeAssoStruct.ItemIndex;
    ListeAsso.Items.Exchange(i_ind1, i_ind1 + 1);
    ListeAssoCh.Items.Exchange(i_ind1, i_ind1 + 1);
    ListeAssoStruct.Items.Exchange(i_ind1, i_ind1 + 1);
    StringGridAsso.Rows[0].Exchange(i_ind1, i_ind1 + 1);
    StringGridAssoFic.Rows[0].Exchange(i_ind1, i_ind1 + 1);
    StringGridAssoStruct.Rows[0].Exchange(i_ind1, i_ind1 + 1);
    ListeAsso.ItemIndex := i_ind1 + 1;
    ListeAssoCh.ItemIndex := i_ind1 + 1;
    ListeAssoStruct.ItemIndex := i_ind1 + 1;
    Maj_Edit3;
end;

procedure TFAssocImport.BitBtn9Click(Sender: TObject);
begin
  if ListeAssoCh.ItemIndex = -1 then Exit;
  if Edit4.Text = '' then
    Edit4.Text := StringGridAssoFic.Cells[0, ListeAssoCh.ItemIndex];
  if ListeFichier.ItemIndex = -1 then Exit;
  if Pos('=', Edit4.Text) = 0 then
      Edit4.Text := Edit4.Text + '=' + StringGridFichier.Cells[0, ListeFichier.ItemIndex]
  else
//  Modif 07/02/02 rapprochement GC - GRC
  begin
    if  Pos(StringGridAssoFic.Cells[0, ListeAssoCh.ItemIndex], Edit4.Text) = 0 then
      Edit4.Text := Edit4.Text+';'+StringGridAssoFic.Cells[0,ListeAssoCh.ItemIndex]+'='+StringGridFichier.Cells[0, ListeFichier.ItemIndex]
    else
      Edit4.Text := Edit4.Text + '+' + StringGridFichier.Cells[0, ListeFichier.ItemIndex];
  end;
//  Fin Modif 07/02/02 rapprochement GC - GRC
end;

procedure TFAssocImport.ListeTableEnter(Sender: TObject);
begin
{    if Sender is TListBox then
        st_ListeTable_Rech := '';  }
end;

procedure TFAssocImport.ListeTableEntree(Sender: TObject; var Key: Word);
begin
{    if Sender = ListeTable then BitBtn1Click(Sender);
    if Sender = ListeFichier then BitBtn4Click(Sender);
    if Sender = ListeAsso then BitBtn2Click(Sender);
    if Sender = ListeAssoCh then BitBtn2Click(Sender);
    if Sender = ListeAssoStruct then BitBtn2Click(Sender); }
end;

procedure TFAssocImport.ListeFichierKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
    i_ind1 : integer;

begin
{    if Sender is TListBox then
    begin
        if (Key = VK_RETURN) then ListeTableEntree(Sender, Key)
        else
        begin
            if (Char(Key) in ['A'..'Z']) or
               (Char(Key) in ['a'..'z']) or
               (Char(Key) = '_') then
            begin
                st_ListeTable_Rech := st_ListeTable_Rech + Char(Key);
                i_ind1 := 0;
                while i_ind1 <= TListBox(Sender).Items.Count - 1 do
                begin
                    if TListBox(Sender).Items.Strings[i_ind1] = '' then
                    begin
                        i_ind1 := -1;
                        break;
                    end;
                    if Pos(st_ListeTable_Rech, TListBox(Sender).Items.Strings[i_ind1]) <> 0 then break;
                    Inc(i_ind1);
                end;
                if i_ind1 > TListBox(Sender).Items.Count - 1 then
                begin
                    Beep;
                    st_ListeTable_Rech := Copy(st_ListeTable_Rech, 0, Length(st_ListeTable_Rech) - 1);
                    i_ind1 := -1;
                end;
                if i_ind1 <> -1 then TListBox(Sender).ItemIndex := i_ind1;
            end;
        end;
    end;    }
end;

procedure TFAssocImport.RadioGroup4Click(Sender: TObject);
var
    i_ind1 : integer;

begin
    ListeFichier.Items.Text := '';
    i_ind1 := 0;
    while StringGridFichier.Cells[RadioGroup4.ItemIndex, i_ind1] <> '' do
    begin
        if i_ind1 = 0 then
            ListeFichier.Items.Text := StringGridFichier.Cells[RadioGroup4.ItemIndex, i_ind1]
        else
            ListeFichier.Items.Add(StringGridFichier.Cells[RadioGroup4.ItemIndex, i_ind1]);
        i_ind1 := Succ(i_ind1);
    end;
end;

procedure TFAssocImport.BitBtn10Click(Sender: TObject);
var
    i_ind1, i_ItemIndexSave : integer;

begin
//  Pas de ligne selectionnee au centre
    if ListeAsso.ItemIndex < 0 then Exit;
    RadioGroup4.ItemIndex := 1;
    ListeDictFichier;
    ListeTable.Items.Text := '';
    i_ind1 := 0;
    while StringGridTable.Cells[1, i_ind1] <> '' do
    begin
        if i_ind1 = 0 then
            ListeTable.Items.Text := StringGridTable.Cells[1, i_ind1]
        else
            ListeTable.Items.Add(StringGridTable.Cells[1, i_ind1]);
        i_ind1 := Succ(i_ind1);
    end;
//  Recupere le nom du champ dans la ligne centrale
    i_ind1 := ListeAsso.ItemIndex;
    i_ItemIndexSave := ListeAsso.ItemIndex;
    FSaisieFormule.ListeTable.Items := ListeTable.Items;
    FSaisieFormule.ListeFichier.Items := ListeFichier.Items;
    RadioGroup4.ItemIndex := 0;
    ListeDictFichier;
    ListeTable.Items.Text := '';
    i_ind1 := 0;
    while StringGridTable.Cells[0, i_ind1] <> '' do
    begin
        if i_ind1 = 0 then
            ListeTable.Items.Text := StringGridTable.Cells[0, i_ind1]
        else
            ListeTable.Items.Add(StringGridTable.Cells[0, i_ind1]);
        i_ind1 := Succ(i_ind1);
    end;
    FSaisieFormule.ListeTableCode.Items := ListeTable.Items;
    FSaisieFormule.ListeFichierCode.Items := ListeFichier.Items;
    FSaisieFormule.Left := (Left + Trunc(Width / 2)) - Trunc(FSaisieFormule.Width / 2);
    FSaisieFormule.Top := (Top + Trunc(Height / 2)) - Trunc(FSaisieFormule.Height / 2);
//    FSaisieFormule.Edit1.Text := Copy(StringGridAsso.Cells[1, i_ind1], 6, 255);
    FSaisieFormule.Edit1.Text := '';
    FSaisieFormule.ShowModal;
    if FSaisieFormule.Edit1.Text <> '' then
    begin
//  Recupere le nom du champ dans la ligne centrale
        i_ind1 := ListeAsso.ItemIndex;
//  Recherche de la ligne selection dans la StringGrid
        if i_ind1 <= StringGridAsso.RowCount then
        begin
            StringGridAsso.Cells[1, i_ind1] := '"calc=' + FSaisieFormule.Edit1.Text + '"';
            StringGridAssoStruct.Cells[1, i_ind1] := '"calc=' + FSaisieFormule.Edit1.Text + '"';
            StringGridAssoFic.Cells[1, i_ind1] := '"calc=' + FSaisieFormule.Edit1.Text + '"';
        end;
        ListeDetailAsso;
        ListeTable.Items.Text := '';
        i_ind1 := 0;
        while StringGridTable.Cells[RadioGroup1.ItemIndex, i_ind1] <> '' do
        begin
            if Pos(StringGridTable.Cells[0, i_ind1], ListeAsso.Items.Text) = 0 then
                if i_ind1 = 0 then
                    ListeTable.Items.Text := StringGridTable.Cells[RadioGroup1.ItemIndex, i_ind1]
                else
                    ListeTable.Items.Add(StringGridTable.Cells[RadioGroup1.ItemIndex, i_ind1]);
            i_ind1 := Succ(i_ind1);
        end;
        ListeAsso.ItemIndex := i_ItemIndexSave;
    end;
end;

procedure TFAssocImport.BitBtn11Click(Sender: TObject);
var
    i_ind1, i_ItemIndexSave : integer;
    st1, st2 : string;
    Table, Fichier, TableCode, FichierCode : string;

begin
//  Pas de ligne selectionnee au centre
    if ListeAsso.ItemIndex < 0 then Exit;
//  Recupere le nom du champ dans la ligne centrale
    i_ind1 := ListeAsso.ItemIndex;
    i_ItemIndexSave := ListeAsso.ItemIndex;
//
    RadioGroup4.ItemIndex := 1;
    ListeDictFichier;
    ListeTable.Items.Text := '';
    i_ind1 := 0;
    while StringGridTable.Cells[1, i_ind1] <> '' do
    begin
        if i_ind1 = 0 then
            ListeTable.Items.Text := StringGridTable.Cells[1, i_ind1]
        else
            ListeTable.Items.Add(StringGridTable.Cells[1, i_ind1]);
        i_ind1 := Succ(i_ind1);
    end;
    Table := ListeTable.Items.Text;
    Fichier := ListeFichier.Items.Text;
    RadioGroup4.ItemIndex := 0;
    ListeDictFichier;
    ListeTable.Items.Text := '';
    i_ind1 := 0;
    while StringGridTable.Cells[0, i_ind1] <> '' do
    begin
        if i_ind1 = 0 then
            ListeTable.Items.Text := StringGridTable.Cells[0, i_ind1]
        else
            ListeTable.Items.Add(StringGridTable.Cells[0, i_ind1]);
        i_ind1 := Succ(i_ind1);
    end;
    TableCode := ListeTable.Items.Text;
    FichierCode := ListeFichier.Items.Text;
    if StringGridAsso.Cells[1, i_ItemIndexSave] = '' then
        st1 := ''
        else
        st1 := StringGridAsso.Cells[1, i_ItemIndexSave];
    st2 := Entree_Condition([Table, Fichier, TableCode, FichierCode, st1], 5);
    if st2 <> '' then
    begin
//  Recupere le nom du champ dans la ligne centrale
        i_ind1 := ListeAsso.ItemIndex;
//        if i_ind1 < 0 then
//            for i_ind1 := 0 to StringGridAsso.RowCount - 1 do
//                if StringGridAsso.Cells[1, i_ind1] = '' then Break;
//  Recherche de la ligne selection dans la StringGrid
        if i_ind1 <= StringGridAsso.RowCount then
        begin
            StringGridAsso.Cells[1, i_ind1] := st2;
            StringGridAssoStruct.Cells[1, i_ind1] := st2;
            StringGridAssoFic.Cells[1, i_ind1] := st2;
        end;
        ListeDetailAsso;
        ListeTable.Items.Text := '';
        i_ind1 := 0;
        while StringGridTable.Cells[RadioGroup1.ItemIndex, i_ind1] <> '' do
        begin
            if i_ind1 = 0 then
                ListeTable.Items.Text := StringGridTable.Cells[RadioGroup1.ItemIndex, i_ind1]
            else
                ListeTable.Items.Add(StringGridTable.Cells[RadioGroup1.ItemIndex, i_ind1]);
            i_ind1 := Succ(i_ind1);
        end;
        ListeAsso.ItemIndex := i_ItemIndexSave;
    end;
end;

end.
