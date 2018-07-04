unit RecupNeg;

//
//   ATTENTION, ce Unit contient des lignes specifiques à destination d'un import
//   dans SQL 7.
//   Elles sont reperées par le commentaire : SPECIFIQUE POUR SQL 7
//

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls,Hent1, HCtrls,ComCtrls, HRichEdt, HRichOLE,UTob, UTom, ParamSoc,
  ExtCtrls, Grids, Db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} ImgList, HStatus, Spin, HPanel, UiUtil,
  Buttons, IniFiles, Math, Ent1, EntGC, VisuLog, UtilArticle, AssistUtil,RecupNegSQL7,
  TntButtons, TntGrids;
//  UTomTiers;

type
  TFRecupNeg = class(TForm)
    BRecup: TButton;
    bVisu: TButton;
    HGrid1: THGrid;
    OpenDialogRep: TOpenDialog;
    TREPERTOIRE: TLabel;
    ListePresent: TListBox;
    ListeATraiter: TListBox;
    bAjouter: THBitBtn;
    bEnlever: THBitBtn;
    Label1: TLabel;
    bTest: TButton;
    ListeAsso: TListBox;
    ListeAssoCh: TListBox;
    ListeAssoStruct: TListBox;
    StringGrid1: TStringGrid;
    StringGrid2: TStringGrid;
    StringGridAsso: TStringGrid;
    StringGridAssoCh: TStringGrid;
    StringGridAssoStruct: TStringGrid;
    Repertoire: TEdit;
    bStop: THBitBtn;
    bAjouterTout: THBitBtn;
    bEnleverTout: THBitBtn;
    StringGridDECHAMPS: TStringGrid;
    StringGrid4: TStringGrid;
    bLog: TButton;
    StringGridFichier: TStringGrid;
    Edit1: TEdit;
    RadioGroup1: TRadioGroup;
    Memo1: TMemo;
    ListeATRacine: TListBox;
    ListeFichiersSeq: TListBox;
    Label2: TLabel;
    RadioGroup3: TRadioGroup;
    WorkingPath: TEdit;
    RGSortie: TRadioGroup;
    ListeAssoCle: TListBox;
    Label3: TLabel;
    Label4: TLabel;
    procedure BRecupClick(Sender: TObject);
    procedure HGrid1RowEnter(Sender: TObject; Ou: Integer;
      var Cancel: Boolean; Chg: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bAjouterClick(Sender: TObject);
    procedure bEnleverClick(Sender: TObject);
    procedure bStopClick(Sender: TObject);
    procedure bAjouterToutClick(Sender: TObject);
    procedure bEnleverToutClick(Sender: TObject);
    procedure ListePresentDblClick(Sender: TObject);
    procedure ListeATraiterDblClick(Sender: TObject);
    procedure bTestClick(Sender: TObject);
    procedure bVisuClick(Sender: TObject);
    procedure bLogClick(Sender: TObject);
   private
    { Déclarations privées }
    Correspond : TStrings;
//    ArtCodeOld : TStrings;
//    ArtCodeNew : TStrings;
    ErreurCommun : TStrings;
    ParamGenCle : TStrings;
    ParamGenData : TStrings;
    DimensionsVal : TStrings;
    DimensionsCle : TStrings;
    TOBBuffer : TOB;
    TobTable : TOB;
//    TomTable : TOM; //PCS
    Tob_DeChamps : Tob;
    Tob_DeCombos : Tob;
    Tob_DeTables : Tob;
    Tob_ChoixCod : Tob;
    Tob_ChoixExt : Tob;
    Tob_Commun   : Tob;
    NumChArray : array [0..1000] of integer;
    CleArray : array [0..1000] of string;
    TOBArray : array [0..1000] of TOB;
    TOBALire : TOB;
    TOBCompteur : TOB;
    Fichier : textfile;
    FicLog : textfile;
    ErreurLog : textfile;
    ErreurDouble : textfile;
//
    FicSeq : textfile;
//
    Corresp : textfile;
//    st_Corresp : string;
    st_enreg : string;
    st_chemin : string;
    st_DateDispo : string;
    Liste_Champs : string;
    Liste_Champs2 : string;
    st_temp1 : string;
    LgCptAux : integer;
    CarBourrage : string;
    l_Arret : Boolean;
    i_LusTotal : Integer;
    Nb_Lus : Integer;
    Nb_Ecrit : Integer;
    NbMax : integer;
    SQL_Insert, SQL_Update, SQL_Champs, SQL_Datas, SQL_Where, SQL_Sep : string;
    NomsTable, NomsFichier : string;
    procedure Charge_Ini;
    procedure Charge_Dict(i_indice : integer; var Traite : Boolean);
    procedure Ferme_Dict;
    procedure Traitement;
    procedure Controle_Ordre_De_Traitement;
    procedure Traiter_Donnees(i_indfic : Integer);
    procedure Traiter_Champ_NEGOCE(NomChampTable : string; NomChampSeq : string; var ValeurSeq : Variant);
    procedure Traiter_Champ_PGI(NomChampTable : string; NomChampSeq : string; var ValeurSeq : Variant);
    procedure Ecrire_Tob(NumChamp : integer; ValeurSeq : Variant);
    procedure Recherche_Tablette(NomChampTable : string; i_ind1 : integer);
    procedure Traiter_Tablette(ChampAssocie, TabLibelle, TabAbrege : string);
    procedure Champ_Tablette(NomChampTable, Valeur : string; i_ind1 : integer);
    procedure Extraire_Champ(i_ind1 : Integer; NomChampTable, st_trav1 : string; var ValeurSeq : Variant;
                             var ChampSeq : Variant);
    function  PresenceComp(Fichier,Champs,Valeurs : String ) : Boolean ;
    procedure Enreg_Log_Erreur;
    function  VerifCorrespondance (i_nb_champs : integer; st_champs : string;
                                   var i_nrow : integer) : boolean;
    procedure ImportEnreg ( TTableRecup : TOB; MAJ : string );
    function  Incremente(NomTable, NomChamp : string; Remplace : Boolean;
                        var ValeurSeq : string) : Integer;
    procedure Traiter_Calcul(NomChampTable : string; var ValeurSeq : Variant);
    function  Evaluer_Si(ValeurSeq, NomChampTable : string) : Variant;
    function  Evaluer_Condition(TSI : TOB; NomChampTable : string) : boolean;
    function  Evaluer_Valeur(TSI : TOB; NomChampTable : string) : variant;
    function  FctPredefinie(st_trav1, NomChampTable : string; var ValeurSeq : Variant) : boolean;
    function  Fct_MAJUSCULE(var Syntaxe : string) : variant;
    function  Fct_COMPTEUR(NomChampTable : string) : variant;
    function  Fct_REPLACECAR(var Syntaxe : string) : Variant;
    function  Fct_CORRESPOND(NomChampTable, NomChampSeq : string) : Variant;
    function  Fct_TABLE(NomChampTable, NomChampSeq : string) : Variant;
    function  Fct_SOUSCHAINE(Syntaxe : string) : Variant;
    function  Fct_SPACE(Syntaxe : string) : Variant;
    function  Fct_INTSTR(Syntaxe : string) : Variant;
    function  Fct_STRINT(Syntaxe : string) : Variant;
    function  Fct_DATE(Syntaxe : string) : Variant;
    function  RecupValeur(Syntaxe : string) : Variant;
    procedure Prepare_Dispo(var traite : boolean);
    procedure GereConvertError(NomChampTable : string; ValeurSeq : Variant);
    procedure GereEngineError(NomTable, SQL : string);
    procedure VerifieSQL(var Sql : string);
  public
    { Déclarations publiques }
  end;
type EConvertError = class(Exception);
type EVariantError = class(Exception);
type EInOutError = class(Exception);
type EDBEngineError = class(Exception);

Procedure RecuperationNeg (WorkingPath : string) ;

implementation

uses HMsgBox;  //PCS

var stDossier, stArrondi : string;

{$R *.DFM}


{----------------------------------------------------------------------------
 Procedure  Point d'entrée du module de récupération des fichiers
            provenant de NEGOCE
-----------------------------------------------------------------------------}
Procedure RecuperationNeg (WorkingPath : string) ;
var
    Fo_Recup : TFRecupNeg;
    Pn_Inside : THPanel ;
begin
    EDBEngineError.Create('');
    EConvertError.Create('');
    EVariantError.Create('');
    Pn_Inside :=  FindInsidePanel;
    Fo_Recup := TFRecupNeg.Create (Pn_Inside);
    Fo_Recup.WorkingPath.Text := WorkingPath;
//    Fo_Recup.Top := Pn_Inside.BoundsRect.Top + Trunc((Pn_Inside.BoundsRect.BottomRight.Y - Fo_Recup.Height) / 2);
//    Fo_Recup.Left := Pn_Inside.BoundsRect.Left + Trunc((Pn_Inside.BoundsRect.BottomRight.X - Fo_Recup.Width) / 2);
    try
        Fo_Recup.ShowModal;
    finally
        Fo_Recup.Free;
    end;
end;

{----------------------------------------------------------------------------
 Procedure d'enregistrement de la TOB dans la table du SGBD
-----------------------------------------------------------------------------}
procedure TFRecupNeg.ImportEnreg ( TTableRecup : TOB; MAJ : string );
var
    vt_trav1 : variant;
    st_trav1, st_Log, CleTable, EnregPGI, OrdreSQL, OrdreSQLCre, OrdreSQLMaj : string;
    i_ind1 : integer;
    EnregExist, BlobPresent : boolean;
    Heure, Minute, Seconde, MSeconde : Word;

begin
BlobPresent := False;
//TobTable.SetAllModifie(False);
//  Modif 06/02/02 rapprochement GC - GRC
//if (MAJ[1] = 'P') or (MAJ[1] = 'M') then
if (MAJ[1] = 'P') or (MAJ[1] = 'M') or (MAJ[2] = 'M') then
//  Fin Modif 06/02/02 rapprochement GC - GRC
    begin
    SQL_Champs := '';
    SQL_Datas  := '';
    SQL_Where  := '';
    SQL_Sep    := '';
    TobTable.ChargeCle1;
    CleTable := TobTable.Cle1;
    for i_ind1 := 0 to StringGridAsso.RowCount - 1 do
        begin
        if StringGridAsso.Cells[0, i_ind1] = '' then Break;
        st_trav1 := StringGridAsso.Cells[0, i_ind1];
        vt_trav1 := TobTable.GetValeur(NumChArray[i_ind1]);
        if Pos('Tablette', st_trav1) = 0 then
            begin
            Ecrire_Datas(RGSortie.ItemIndex, st_trav1, TobTable.NumTable, NumChArray[i_ind1],
                         vt_trav1, SQL_Champs, SQL_Datas, SQL_Sep, BlobPresent);
            if V_PGI.DEChamps[TobTable.NumTable, NumChArray[i_ind1]].Tipe = 'BLOB' then
                begin
                vt_trav1 := TobTable.GetValue(st_trav1);
                TobTable.PutValue(st_trav1, '');
                TobTable.PutValue(st_trav1, vt_trav1);
                end;
            end;
        end;
        SQL_Sep := '';
        OrdreSQLMaj := SQL_Update;
//  Modif 07/02/02 rapprochement GC - GRC
        if BlobPresent and (trim (SQL_Champs) = '') then
        begin
          TobTable.UpdateDB;
          exit;
        end else
//  Fin Modif 07/02/02 rapprochement GC - GRC
        while SQL_Champs <> '' do
            begin
            OrdreSQLMaj := OrdreSQLMaj + SQL_Sep + Trim(ReadTokenStV(SQL_Champs)) +
                          '=' + Trim(ReadTokenStV(SQL_Datas));
            SQL_Sep := ',';
            end;

        OrdreSQLMaj := OrdreSQLMaj + ' where ' + CleTable;
    end;
if (MAJ[1] = 'C') or (MAJ[2] = 'C') then
    begin
    SQL_Champs := '';
    SQL_Datas  := '';
    SQL_Where  := '';
    SQL_Sep    := '';
    for i_ind1 := 1 to TobTable.NbChamps do
        if V_PGI.DEChamps[TobTable.NumTable, i_ind1].Tipe <> 'BLOB' then
            begin
            st_trav1 := TobTable.GetNomChamp(i_ind1);
            vt_trav1 := TobTable.GetValeur(i_ind1);
            Ecrire_Datas(RGSortie.ItemIndex, st_trav1, TobTable.NumTable, i_ind1,
                         vt_trav1, SQL_Champs, SQL_Datas, SQL_Sep, BlobPresent);
            end;
        OrdreSQLCre := SQL_Insert + ' (' + SQL_Champs + ') values (' + SQL_Datas + ')';
    end;

//
//   ATTENTION : SPECIFIQUE POUR SQL 7
//
if RGSortie.ItemIndex = 0 then
    begin
    try
    TobTable.ChargeCle1;
//  Modif 06/02/02 rapprochement GC - GRC
//    EnregExist := TobTable.existDB;
//    if (MAJ[1] = 'C') then
    if (MAJ[2] = 'C') then  EnregExist := TobTable.existDB;

    if (MAJ[1] = 'C') and (MAJ[2] = 'C')  then
//  Fin Modif 06/02/02 rapprochement GC - GRC
        if EnregExist then
            writeln (ErreurDouble,'Clé double ='+ TobTable.Cle1)
            else
            begin
            OrdreSQL := OrdreSQLCre;
            VerifieSQL(OrdreSQLCre);
            ExecuteSQL(OrdreSQLCre);
            end
        else
        begin
        if (MAJ[1] = 'P') then
            begin
            OrdreSQL := OrdreSQLMaj;
            VerifieSQL(OrdreSQLMaj);
            ExecuteSQL(OrdreSQLMaj);
            end
            else
            begin
            if (MAJ[2] = 'M') or EnregExist then
                begin
                OrdreSQL := OrdreSQLMaj;
                VerifieSQL(OrdreSQLMaj);
                ExecuteSQL(OrdreSQLMaj);
                end
                else
                begin
                OrdreSQL := OrdreSQLCre;
                VerifieSQL(OrdreSQLCre);
                ExecuteSQL(OrdreSQLCre);
                end
            end;
        end;
    except
        on EDBEngineError do GereEngineError(TobTable.NomTable, OrdreSQL);
    end;
//    Memo1.Lines.Text := VarAsType(TobTable.GetValue(TableToPrefixe(TobTable.Nomtable)+'_BLOCNOTE'), varString);

    if BlobPresent then TobTable.UpdateDB;
    if (Nb_Ecrit mod 500) = 0 then
        begin
        Try
            CommitTrans;
            DecodeTime(Now, Heure, Minute, Seconde, MSeconde);
            st_log := 'Ecriture de 500 enregs dans la base : ' + IntToStr(Heure) + ':' + IntToStr(Minute) +
                      ':' + IntToStr(Seconde) + '.' + IntToStr(MSeconde);
            writeln (FicLog, st_log);
        Except
            RollBack ;
        end;
        BeginTrans;
        end;
    end
    else
    begin
        EnregPGI := SQL_Datas;
        writeln (FicSeq, EnregPGI);
        flush (FicSeq);
    end;
end ;

{----------------------------------------------------------------------------
 Procedure d'action sur le bouton lançant la récupération de fichiers
-----------------------------------------------------------------------------}
procedure TFRecupNeg.BRecupClick (Sender: TObject);
begin
    Controle_Ordre_De_Traitement;
    NbMax := High(NbMax);
    Traitement;
End;

//----------------------------------------------------------------------------
procedure TFRecupNeg.bTestClick(Sender: TObject);
begin
    Controle_Ordre_De_Traitement;
    NbMax := StrToInt(RadioGroup3.Items[RadioGroup3.ItemIndex]);
    Traitement;
end;

//----------------------------------------------------------------------------
procedure TFRecupNeg.Traitement;
var
    i_ind1, i_ind2 : integer;
    st_log, st_enreg, st_trav1, st_trav2 : string;
    Traite : Boolean;
    Heure, Minute, Seconde, MSeconde : Word;
    Q1 : TQuery;

begin
//    MemCheckLogFileName := 'D:\Pgi\PgiS5\Fichiers\Memcheq.log';
//    MemChk;
    Correspond.Text := '';
    if FileExists (WorkingPath.Text + 'Correspondances') then
        begin
        AssignFile(Corresp, WorkingPath.Text + 'Correspondances');
        Reset (Corresp);
        readln (Corresp,st_enreg);
        while not EOF(Corresp) do
        Begin
            Correspond.Add(st_enreg);
            readln (Corresp,st_enreg);
        end;
        CloseFile(Corresp);
        end;
    AssignFile(Corresp, WorkingPath.Text + 'Correspondances');
    if Correspond.Count < 1 then
        Rewrite (Corresp)
        else
        Append (Corresp);
//
    Tob_DeTables := TOB.Create('', nil, -1);
    Q1 := OpenSQL('Select * from DETABLES', True,-1,'',true);
    if Q1.Eof then Exit;
    Tob_DeTables.LoadDetailDB('DETABLES', '', '', Q1, False);
    Ferme(Q1);
    Tob_DeChamps := TOB.Create('', nil, -1);
    Q1 := OpenSQL('Select * from DECHAMPS', True,-1,'',true);
    if Q1.Eof then Exit;
    Tob_DeChamps.LoadDetailDB('DECHAMPS', '', '', Q1, False);
    Ferme(Q1);
    Tob_DeCombos := TOB.Create('', nil, -1);
    Q1 := OpenSQL('Select * from DECOMBOS', True,-1,'',true);
    if Q1.Eof then Exit;
    Tob_DeCombos.LoadDetailDB('DECOMBOS', '', '', Q1, False);
    Ferme(Q1);
    Tob_ChoixCod := TOB.Create('', nil, -1);
    Q1 := OpenSQL('Select * from CHOIXCOD', True,-1,'',true);
//    if Q1.Eof then Exit;
    Tob_ChoixCod.LoadDetailDB('CHOIXCOD', '', '', Q1, False);
    Ferme(Q1);
    Tob_ChoixExt := TOB.Create('', nil, -1);
    Q1 := OpenSQL('Select * from CHOIXEXT', True,-1,'',true);
    Tob_ChoixExt.LoadDetailDB('CHOIXEXT', '', '', Q1, False);
    Ferme(Q1);
    Tob_Commun := TOB.Create('', nil, -1);
    Q1 := OpenSQL('Select * from COMMUN', True,-1,'',true);
    if Q1.Eof then Exit;
    Tob_Commun.LoadDetailDB('COMMUN', '', '', Q1, False);
    Ferme(Q1);
    LgCptAux := GetParamSoc('SO_LGCPTEAUX');
    CarBourrage := GetParamSoc('SO_BOURREAUX');
    ErreurCommun := TstringList.Create;
    ListeATRacine.Items.Text := '';
    i_LusTotal := 0;
    for i_ind1 := 0 to ListeATraiter.Items.Count - 1 do
    begin
        bStop.Enabled := True;
        TOBCompteur := TOB.Create('', nil, -1);
        ListeATraiter.ItemIndex := i_ind1;
        ListeATraiter.Selected[i_ind1] := True;
        Repertoire.Text := ListeATraiter.Items.Strings[i_ind1];
        for i_ind2 := 0 to StringGrid1.RowCount - 1 do
            StringGrid1.Rows[i_ind2].Clear;
        AssignFile (FicLog, WorkingPath.Text + Repertoire.Text + '.log');
        Rewrite (FicLog);
        st_log := '****** Début du traitement ( ' + Repertoire.Text + ' ) ******';
        writeln (FicLog, st_log);
        flush (FicLog);
        AssignFile (ErreurLog, WorkingPath.Text + Repertoire.Text + '_erreur.log');
        Rewrite (ErreurLog);
        AssignFile (ErreurDouble, WorkingPath.Text + Repertoire.Text + '_doubles.log');
        Rewrite (ErreurDouble);
        st_log := '****** Erreurs dans le traitement de ' + Repertoire.Text + ' ******';
        writeln (ErreurLog, st_log);
        flush (ErreurLog);
        Charge_Ini;
        for i_ind2 := 0 to StringGrid1.ColCount - 1 do
        begin
            if StringGrid1.Cells[i_ind2, 0] = '' then Break;
            Charge_Dict(i_ind2, Traite);
            if not Traite then Continue;
//
//   ATTENTION : SPECIFIQUE POUR SQL 7
//
            if RGSortie.ItemIndex <> 0 then
                begin
                AssignFile (FicSeq, WorkingPath.Text + Repertoire.Text + '.PGI');
                Rewrite (FicSeq);
                end;

            st_trav1 := ParamGenData.Strings[ParamGenCle.IndexOf('DelFic')];
            st_trav1 := Copy(st_trav1, Pos('=', st_trav1) + 1, 255);
            st_trav2 := ParamGenData.Strings[ParamGenCle.IndexOf('ModeCreation')];
            st_trav2 := Copy(st_trav2, Pos('=', st_trav2) + 1, 255);
            if (st_trav1 = 'O') and (st_trav2 = 'O')  and (i_ind2 = 0) then
//  mode creation et premier fichier de la liste, vide la table et les index
                begin
                if ((uppercase(Repertoire.Text) <> 'CHOIXCOD') and (uppercase(Repertoire.Text) <> 'COMMUN')) then
                ExecuteSQL ('delete from ' + Repertoire.Text);
                end;

            Traiter_Donnees(i_ind2);
            Enreg_Log_Erreur;
            DecodeTime(Now, Heure, Minute, Seconde, MSeconde);
            st_log := 'Heure de fin : ' + IntToStr(Heure) + ':' + IntToStr(Minute) +
                      ':' + IntToStr(Seconde) + '.' + IntToStr(MSeconde);
            writeln (FicLog, st_log);
            flush (FicLog);
            if l_Arret then Break;
        end;
        st_log := '****** Fin du traitement ( ' + Repertoire.Text + ' ) ******';
        writeln (FicLog, st_log);
        flush (FicLog);
//
//   ATTENTION : SPECIFIQUE POUR SQL 7
//
        if RGSortie.ItemIndex <> 0 then CloseFile(FicSeq);
//
        Ferme_Dict;
        TOBCompteur.Free;
        Tob_ChoixCod.InsertOrUpdateDB;
        Tob_ChoixExt.InsertOrUpdateDB;
        if l_Arret then Break;
    end;
    if i_LusTotal > 0 then bVisu.Enabled := True;
//  if i_LusTotal > 0 then bLog.Enabled := True;
    bLog.Enabled := True;
    ErreurCommun.Free;
    Tob_DeTables.Free;
    Tob_Dechamps.Free;
    Tob_Decombos.Free;
    Tob_ChoixCod.Free;
    Tob_ChoixExt.Free;
    Tob_Commun.Free;
    bStop.Enabled := False;
    CloseFile (Corresp);
End;

//----------------------------------------------------------------------------
procedure TFRecupNeg.bVisuClick(Sender: TObject);
var
    T1 : TOB;
    FicIni : TIniFile;
    st_trav1, st_racine : string;
    i_ind1 : integer;

begin
    Repertoire.Text := ListeATraiter.Items.Strings[ListeATraiter.ItemIndex];
    st_racine := ListeATRacine.Items.Strings[ListeATraiter.ItemIndex];
    st_Chemin := WorkingPath.Text;
    FicIni := TIniFile.Create(st_chemin + Repertoire.Text + '.ini');
    ListeAssoCh.Items.Text := '';
    FicIni.ReadSectionValues('Champs_' + st_racine,ListeAssoCh.Items);
    Liste_Champs := '';
    Liste_Champs2 := '';
    i_ind1 := 0;
    while i_ind1 <= ListeAssoCh.Items.Count - 1 do
    begin
        st_trav1 := ListeAssoCh.Items.Strings[i_ind1];
        if Copy(st_trav1, 0, 8) <> 'Tablette' then
        begin
        if i_ind1 = 0 then
            Liste_Champs := Copy(st_trav1, 0, Pos('=', st_trav1) -1)
        else
            Liste_Champs := Liste_Champs + ';' + Copy(st_trav1, 0, Pos('=', st_trav1) -1);
        if i_ind1 < 2 then
            if i_ind1 = 0 then
                Liste_Champs2 := Copy(st_trav1, 0, Pos('=', st_trav1) -1)
            else
                Liste_Champs2 := Liste_Champs2 + ';' + Copy(st_trav1, 0, Pos('=', st_trav1) -1);
        end;
        inc(i_ind1);
    end;
    T1 := TOB.Create(Repertoire.Text, NIL, -1);
    Hgrid1.VidePile(true) ;
    T1.LoadDetailDB(Repertoire.Text, '', '',NIL,True,True);
    T1.PutGridDetail(HGrid1, TRUE, TRUE, Liste_Champs);
    FicIni.Free;
    T1.Free;
end;

//----------------------------------------------------------------------------
procedure TFRecupNeg.bLogClick(Sender: TObject);
    var i_ind1 : integer;
        fdoubles : string;
        SearchRec: TSearchRec;
begin
    FVisuLog := TFVisuLog.Create(FVisuLog);
    FVisuLog.ListBox1.Items.Text := ListeATraiter.Items.Text;
    FVisuLog.ListBox1.ItemIndex := ListeATraiter.ItemIndex;
//{$IFDEF GRC} Modif 06/02/02 rapprochement GC - GRCAB
    for i_ind1 := 0 to ListeATraiter.Items.Count - 1 do
    begin
      fdoubles := ListeATraiter.Items.Strings[i_ind1]+'_doubles';
      if ( (FindFirst(st_chemin +fdoubles+ '.log', faAnyFile, SearchRec) = 0)
         and (SearchRec.Size > 0) ) then
        FVisuLog.ListBox1.Items.append(fdoubles);
    end;
//{$ENDIF} Fin Modif 06/02/02 rapprochement GC - GRCAB
    FVisuLog.CheminLog.text := st_chemin;
    FVisuLog.ShowModal;
    FVisuLog.free;
    FVisuLog :=Nil;
end;

//----------------------------------------------------------------------------
procedure TFRecupNeg.HGrid1RowEnter(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
  var T2 : TOB;
begin
   T2 := TOB(Hgrid1.objects[0,Hgrid1.row]);
   T2.PutEcran(self);
end;

//----------------------------------------------------------------------------
procedure TFRecupNeg.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    If isInside (Self) Then
    Begin
        //ArtCodeOld.Free;
        //ArtCodeNew.Free;
        Action := caFree;
    end;
end;

//**************************************************************************
//**************************************************************************
//**************************************************************************

procedure TFRecupNeg.FormCreate(Sender: TObject);
begin
    Correspond := TStringList.Create;
    ParamGenCle := TStringList.Create;
    ParamGenData := TStringList.Create;
    DimensionsVal := TStringList.Create;
    DimensionsCle := TStringList.Create;
end;

//----------------------------------------------------------------------------
procedure TFRecupNeg.FormShow(Sender: TObject);
var
    SearchRec: TSearchRec;
    FicIni : TIniFile;
    i_ind1 : integer;

begin
    Label3.Caption := '';
    st_Chemin := WorkingPath.Text;
    FicIni := TIniFile.Create(st_chemin + 'ListeFic');
    stDossier := FicIni.ReadString('DossierActif', 'Actif', '');
    if stDossier <> '' then stDossier := '_' + stDossier;
    FicIni.ReadSection('General' + stDossier, ParamGenCle);
    for i_ind1 := 0 to ParamGenCle.Count - 1 do
        ParamGenData.Append(FicIni.ReadString('General' + stDossier, ParamGenCle.Strings[i_ind1], ''));
    stArrondi := FicIni.ReadString('General' + stDossier, 'Arrondi', '');
    FicIni.Free;
    if FindFirst(st_chemin + '*.ini', faAnyFile, SearchRec) = 0 then
    begin
        ListePresent.Items.Text := Copy(SearchRec.Name, 0, Pos('.', SearchRec.Name) - 1);
        while FindNext(SearchRec) = 0 do
            ListePresent.Items.Append(Copy(SearchRec.Name, 0, Pos('.', SearchRec.Name) - 1));
    end;
    bRecup.Enabled := False;
    bTest.Enabled := False;
    bVisu.Enabled := False;
    bLog.Enabled := False;
    bEnlever.Enabled := False;
    bEnleverTout.Enabled := False;
    bStop.Enabled := False;
end;

//----------------------------------------------------------------------------
procedure TFRecupNeg.bAjouterClick(Sender: TObject);
var
    i_ind1 : integer;
    st_trav1 : string;

begin
    for i_ind1 := 0 to ListePresent.Items.Count - 1 do
    begin
        if ListePresent.Selected[i_ind1] then
        begin
            st_trav1 := ListePresent.Items.Strings[i_ind1];
            ListeATraiter.Items.Append(st_trav1);
        end;
    end;
    i_ind1 := ListePresent.Items.Count - 1;
    while i_ind1 >= 0 do
    begin
        if ListePresent.Selected[i_ind1] then
        begin
            ListePresent.Items.Delete(i_ind1);
        end;
        i_ind1 := i_ind1 - 1;
    end;
    if ListePresent.Items.Count = 0 then bAjouter.Enabled := False;
    if ListePresent.Items.Count = 0 then bAjouterTout.Enabled := False;
    if ListeATraiter.Items.Count > 0 then
    begin
        bRecup.Enabled := True;
        bTest.Enabled := True;
        bEnlever.Enabled := True;
        bEnleverTout.Enabled := True;
    end;
end;

//----------------------------------------------------------------------------
procedure TFRecupNeg.bEnleverClick(Sender: TObject);
var
    i_ind1 : integer;
    st_trav1 : string;

begin
    for i_ind1 := 0 to ListeATraiter.Items.Count - 1 do
    begin
        if ListeATraiter.Selected[i_ind1] then
        begin
            st_trav1 := ListeATraiter.Items.Strings[ListeATraiter.ItemIndex];
            ListePresent.Items.Append(st_trav1);
        end;
    end;
    i_ind1 := ListeATraiter.Items.Count - 1;
    while i_ind1 >= 0 do
    begin
        if ListeATraiter.Selected[i_ind1] then
        begin
            ListeATraiter.Items.Delete(ListeATraiter.ItemIndex);
        end;
        i_ind1 := i_ind1 - 1;
    end;
    if ListePresent.Items.Count > 0 then bAjouter.Enabled := True;
    if ListePresent.Items.Count > 0 then bAjouterTout.Enabled := True;
    if ListeATraiter.Items.Count = 0 then
    begin
        bRecup.Enabled := False;
        bTest.Enabled := False;
        bVisu.Enabled := False;
        bLog.Enabled := False;
        bEnlever.Enabled := False;
        bEnleverTout.Enabled := False;
    end;
end;

//----------------------------------------------------------------------------
procedure TFRecupNeg.bStopClick(Sender: TObject);
begin
    l_Arret := True;
end;

//----------------------------------------------------------------------------
procedure TFRecupNeg.bAjouterToutClick(Sender: TObject);
begin
    ListeATraiter.Items.Text := ListePresent.Items.Text;
    ListePresent.Items.Text := '';
    bRecup.Enabled := True;
    bTest.Enabled := True;
    bEnlever.Enabled := True;
    bEnleverTout.Enabled := True;
end;

//----------------------------------------------------------------------------
procedure TFRecupNeg.bEnleverToutClick(Sender: TObject);
begin
    ListePresent.Items.Text := ListeATraiter.Items.Text;
    ListeATraiter.Items.Text := '';
    bRecup.Enabled := False;
    bTest.Enabled := False;
    bVisu.Enabled := False;
    bLog.Enabled := False;
    bEnlever.Enabled := False;
    bEnleverTout.Enabled := False;
end;

//----------------------------------------------------------------------------
procedure TFRecupNeg.ListePresentDblClick(Sender: TObject);
begin
    bAjouterClick(Sender);
end;

//----------------------------------------------------------------------------
procedure TFRecupNeg.ListeATraiterDblClick(Sender: TObject);
begin
    bEnleverClick(Sender);
end;

//----------------------------------------------------------------------------
procedure TFRecupNeg.Charge_Ini;
var
    st_trav1, st_file : string;
    i_ind1 : integer;
    FicIni : TIniFile;

begin
    if Repertoire.Text = '' then Exit;
    st_Chemin := WorkingPath.Text;
    st_file := st_Chemin + Repertoire.Text;
    if not FileExists (st_file + '.ini') then
    begin
        if FileExists (st_file + '.def') then
        begin
            CopyFile (PChar (st_file + '.def'), PChar (st_file + '.ini'), False);
        end;
    end;
    if FileExists(st_Chemin + Repertoire.Text + '.ini') then
    begin
        FicIni := TIniFile.Create(st_Chemin + Repertoire.Text + '.ini');

        ListeAsso.Items.Text := '';
        FicIni.ReadSectionValues ('Fichier', ListeAsso.Items);
        ListeFichiersSeq.Items.Clear;
        i_ind1 := 0;
        while i_ind1 <= ListeAsso.Items.Count - 1 do
        begin
            st_trav1 := ListeAsso.Items.Strings [i_ind1];
            st_trav1 := Copy(st_trav1, Pos('=', st_trav1) + 1, 255);
            while Pos('\', st_trav1) <> 0 do ReadTokenPipe(st_trav1, '\');
            if i_ind1 = 0 then
                ListeFichiersSeq.Items.Text := ReadTokenSt(st_trav1)
                else
                ListeFichiersSeq.Items.Add(ReadTokenSt(st_trav1));
            st_trav1 := ListeAsso.Items.Strings [i_ind1];
            st_trav1 := Copy(st_trav1, Pos('=', st_trav1) + 1, 255);
            StringGrid1.Cells[i_ind1, 0] := Copy(st_trav1, 0, Pos(';',st_trav1) - 1);
            StringGrid2.Cells[i_ind1, 0] := Copy(st_trav1, Pos(';',st_trav1) + 1, 255);
            if i_ind1 = 0 then
            begin
                st_trav1 := Copy(st_trav1, Pos(';', st_trav1) + 1, 255);
                st_trav1 := Copy(st_trav1, Pos(';', st_trav1) + 1, 255);
                ListeATracine.Items.Append(st_trav1);
            end;
            i_ind1 := i_ind1 + 1;
        end;
        if ListeAsso.Items.Count > 1 then
            begin
            RadioGroup3.Visible := False;
            RGSortie.Visible := False;
            ListeFichiersSeq.Visible := True;
            Label2.Visible := True;
            end
            else
            begin
            RadioGroup3.Visible := True;
            RGSortie.Visible := True;
            ListeFichiersSeq.Visible := False;
            Label2.Visible := False;
            end;
        ListeAsso.Items.Text := '';
        ListeAssoCh.Items.Text := '';
        ListeAssoStruct.Items.Text := '';

        TobTable := TOB.CREATE (Repertoire.Text, NIL, -1);
//        TomTable := CreateTOM(Repertoire.Text, Nil, False);


//********************UNIQUEMENT POUR LES TESTS********************************
//        ExecuteSQL ('Delete from CHOIXCOD');
//*****************************************************************************



    end
    else
    begin
        writeln (FicLog, 'Pas de fichier INI trouvé pour ' + Repertoire.Text);
        flush (FicLog);
        Exit;
    end;
end;

//----------------------------------------------------------------------------
procedure TFRecupNeg.Charge_Dict(i_indice : integer; var Traite : Boolean);
var
    Fichier: textfile;
    st_enreg, st_trav1, st_trav2, st_file, st_racine, st_log,st_MAJ : string;
    i_ind1, i_ind2, i_indtext, i_nrow : integer;
    i_nb_ch_text : integer;
    b_Trouve : boolean;
    FicIni : TIniFile;
    Heure, Minute, Seconde, MSeconde : Word;

begin
    Traite := True;
    st_Chemin := WorkingPath.Text;
    st_file := StringGrid1.Cells[i_indice, 0];
    for i_ind1 := 0 to StringGridFichier.RowCount - 1 do
    begin
        StringGridFichier.Rows[i_ind1].Clear;
        StringGridDECHAMPS.Rows[i_ind1].Clear;
    end;
//  Modif 07/02/02 rapprochement GC - GRC
    for i_ind1 := 0 to StringGridAsso.RowCount - 1 do
    begin
        StringGridAsso.Rows[i_ind1].Clear;
        StringGridAssoCh.Rows[i_ind1].Clear;
        StringGridAssoStruct.Rows[i_ind1].Clear;
    end;

    //  Ouverture dictionnaire fichier texte
    if not FileExists (st_file + '.des') then
    begin
      st_log := 'Erreur d''accès aux fichiers de description : ' + st_file + '.des' ;
      PGIBox(st_log,'Import impossible');
      writeln (FicLog, st_log);
      Traite := False;
      exit;
    end ;
//  Fin Modif 07/02/02 rapprochement GC - GRC
    AssignFile(Fichier, st_file + '.des');
    Reset(Fichier);
    //  Chargement de la liste des champs du fichier texte
    readln(Fichier, st_enreg);
    st_racine := Copy(st_enreg, 0, 3);
    i_ind1 := 0;
    while st_enreg <> '' do
    begin
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

    for i_ind2 := 0 to ParamGenCle.Count - 1 do
        begin
        StringGridFichier.Cells[0, i_ind1] := st_racine + 'PARAM.' + ParamGenCle.Strings[i_ind2];
        StringGridFichier.Cells[1, i_ind1] := st_racine + 'PARAM.' + ParamGenCle.Strings[i_ind2];
        StringGridFichier.Cells[2, i_ind1] := ParamGenData.Strings[i_ind2];
        i_ind1 := Succ(i_ind1);
        end;

    i_nb_ch_text := i_ind1;

    //  Chargement de la version precedente eventuelle
    FicIni := TIniFile.Create(st_chemin + Repertoire.Text + '.ini');

    ListeAsso.Items.Text := '';
    ListeAssoCh.Items.Text := '';
    ListeAssoStruct.Items.Text := '';
//  Modif 06/02/02 rapprochement GC - GRC
    st_MAJ := Copy(StringGrid2.Cells[i_indice, 0], 0,
                          Pos(';', StringGrid2.Cells[i_indice, 0]) - 1);
    if ((st_racine = 'PRO') or ((Edit1.text = 'TIERSCOMPL') and (st_racine = 'GRI'))) then
        st_racine := st_racine + st_MAJ;
//  Fin Modif 06/02/02 rapprochement GC - GRC
    Edit1.Text := FicIni.ReadString('Cle_' + st_racine, 'Cle', '');
    FicIni.ReadSectionValues('Champs_' + st_racine,ListeAssoCh.Items);
    if ListeAssoCh.Items.Count < 1 then
    begin
        Traite := False;       
        Exit;
    end;
    FicIni.ReadSectionValues('Designation_' + st_racine,ListeAsso.Items);
    FicIni.ReadSectionValues('Offset_' + st_racine,ListeAssoStruct.Items);

    writeln (FicLog, 'Fichier de données lu : ' + st_file);
    flush (FicLog);
    DecodeTime(Now, Heure, Minute, Seconde, MSeconde);
    st_log := 'Heure de début : ' + IntToStr(Heure) + ':' + IntToStr(Minute) +
              ':' + IntToStr(Seconde) + '.' + IntToStr(MSeconde);
    writeln (FicLog, st_log);
    flush (FicLog);

    i_ind1 := 0;
    i_indtext := 0;
    while i_indtext <= ListeAssoCh.Items.Count - 1 do
    begin
        b_trouve := True; // valeur initiale
        // mise à jour champs et Offset
        st_trav1 := ListeAssoCh.Items.Strings[i_indtext];
        st_trav2 := ListeAssoStruct.Items.Strings[i_indtext];
        StringGridAssoCh.Cells[0, i_ind1] := Copy(st_trav1, 0, Pos('=',st_trav1) - 1);
        StringGridAssoStruct.Cells[0, i_ind1] := Copy(st_trav2, 0, Pos('=',st_trav2) - 1);
//  on ne fait rien de plus sur un SI eventuel
        if Pos('{SI', st_trav1) <> 0 then
            begin
            StringGridAsso.Cells[0, i_ind1] := Copy(st_trav1, 0, Pos('=',st_trav1) - 1);
            st_trav1 := Copy(st_trav1, Pos('=',st_trav1) + 1, 255);
            st_trav2 := Copy(st_trav2, Pos('(',st_trav2) + 1, 255);
            StringGridAsso.Cells[1, i_ind1] := st_trav1;
            StringGridAssoCh.Cells[1, i_ind1] := st_trav1;
            StringGridAssoStruct.Cells[1, i_ind1] := st_trav2;
            Inc(i_ind1);
            Inc(i_indtext);
            Continue;
            end;
//
        st_trav1 := Copy(st_trav1, Pos('=',st_trav1) + 1, 255);
        st_trav2 := Copy(st_trav2, Pos('(',st_trav2) + 1, 255);
        i_ind2 := 1;
        while Pos('+', st_trav1) <> 0 do
        begin
        // Verif code champs avec fichier texte
            b_trouve := VerifCorrespondance (i_nb_ch_text,
                                Copy(st_trav1, 0, Pos('+',st_trav1) - 1), i_nrow);
            if b_trouve = true then
            begin
                if Copy(st_trav1, 2, 4) = 'calc' then
                begin
                    StringGridAssoCh.Cells[i_ind2, i_ind1] := st_trav1;
                    StringGridAssoStruct.Cells[i_ind2, i_ind1] := st_trav2;
                    Break;
                end
                else
                begin
                    StringGridAssoCh.Cells[i_ind2, i_ind1] := Copy(st_trav1, 0, Pos('+',st_trav1) - 1);
                    st_trav1 := Copy(st_trav1, Pos('+',st_trav1) + 1, 255);
                    if i_nrow = -1 then
                    begin
                        StringGridAssoStruct.Cells[i_ind2, i_ind1] := Copy(st_trav2, 0, Pos(')',st_trav2) - 1);
                    end
                    else
                    begin
                        StringGridAssoStruct.Cells[i_ind2, i_ind1] := StringGridFichier.Cells[2, i_nrow];
                    end;
                    st_trav2 := Copy(st_trav2, Pos('(',st_trav2) + 1, 255);
                    i_ind2 := i_ind2 + 1;
                end;
            end
            else
            begin
                break;
            end;
        end;
        if b_trouve = true then
        begin
        // Verif code champs avec fichier texte
            b_trouve := VerifCorrespondance (i_nb_ch_text, st_trav1, i_nrow);
            if b_trouve = False then
            begin
                ListeAsso.Items.Delete(i_indtext);
                ListeAssoCh.Items.Delete(i_indtext);
                ListeAssoStruct.Items.Delete(i_indtext);
                i_indtext := i_indtext + 1;
                continue;
            end;
            StringGridAssoCh.Cells[i_ind2, i_ind1] := st_trav1;
            if i_nrow = -1 then
            begin
                StringGridAssoStruct.Cells[i_ind2, i_ind1] := Copy(st_trav2, 0, Pos(')',st_trav2) - 1);
            end
            else
            begin
                StringGridAssoStruct.Cells[i_ind2, i_ind1] := StringGridFichier.Cells[2, i_nrow];
            end;

            st_trav1 := ListeAsso.Items.Strings[i_indtext];
            StringGridAsso.Cells[0, i_ind1] := Copy(st_trav1, 0, Pos('=',st_trav1) - 1);
            st_trav1 := Copy(st_trav1, Pos('=',st_trav1) + 1, 255);
            i_ind2 := 1;
            while Pos('+', st_trav1) <> 0 do
            begin
                StringGridAsso.Cells[i_ind2, i_ind1] := Copy(st_trav1, 0, Pos('+',st_trav1) - 1);
                st_trav1 := Copy(st_trav1, Pos('+',st_trav1) + 1, 255);
                i_ind2 := i_ind2 + 1;
            end;
            StringGridAsso.Cells[i_ind2, i_ind1] := st_trav1;
//  recherche dans dechamps du type, du prefixe et du code pour les combos
            StringGridDECHAMPS.Cells[0, i_ind1] := StringGridAsso.Cells[0, i_ind1];
//{$IFNDEF GRC} Modif 06/02/02 rapprochement GC - GRCAB
            Recherche_Tablette(StringGridAsso.Cells[0, i_ind1], i_ind1);
//{$ENDIF} Modif 06/02/02 rapprochement GC - GRCAB
            i_ind1 := i_ind1 + 1;
        end;
        i_indtext := i_indtext + 1;
    end;
    FicIni.Free;
    CloseFile(Fichier);
end;

//----------------------------------------------------------------------------------
function TFRecupNeg.VerifCorrespondance (i_nb_champs : integer;
                                         st_champs : string;
                                         var i_nrow : integer) : boolean;
var
    i_ind : integer;
    b_trouve : boolean;
begin
    b_trouve := False;
    i_nrow := 0;
    if (st_champs[1] = '"') or (st_champs[1] = '{') then
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
procedure TFRecupNeg.Ferme_Dict;
begin
    CloseFile (FicLog);
    CloseFile (ErreurLog);
    CloseFile (ErreurDouble);
    TobTable.Free;
//    TomTable.Free;
end;

//----------------------------------------------------------------------------
procedure TFRecupNeg.Traiter_Donnees(i_indfic : Integer);
var
    st_trav1, st_trav2,st_trav3, st_racine, st_MAJ, st_cletable : string;
    NomChampTable, NomChampSeq, TabLibelle, TabAbrege, TypeChamp, ClePrec : string;
    ValeurSeq, ChampSeq : Variant;
    i_ind1, i_ind2, i_sizefic : integer;
    i_decimale : integer;
    i_PInt, i_PDec, i_Sens : integer;
    Searchrec : TSearchRec;
    l_Suivant, Existe_Tablette : boolean;

begin
    l_Arret := False;
//  determination du radical des champs a traiter
    st_trav1 := StringGrid1.Cells[i_indfic, 0] + '.des';
    AssignFile (Fichier, st_trav1);
    Reset (Fichier);  // equivalent open file
    readln (Fichier, st_enreg);
    st_racine := Copy(StringGrid2.Cells[i_indfic, 0],
                          Pos(';', StringGrid2.Cells[i_indfic, 0]) + 1, 255);
//  Modif 06/02/02 rapprochement GC - GRC
    st_racine := Copy (st_racine,1,3);
//  Fin Modif 06/02/02 rapprochement GC - GRC
    st_MAJ := Copy(StringGrid2.Cells[i_indfic, 0], 0,
                          Pos(';', StringGrid2.Cells[i_indfic, 0]) - 1);
    st_trav1 := ParamGenData.Strings[ParamGenCle.IndexOf('ModeCreation')];
    st_trav1 := Copy(st_trav1, Pos('=', st_trav1) + 1, 255);
    if st_trav1 = 'O' then st_trav1 := 'C';
    st_MAJ := st_trav1 + st_MAJ;
//  Creation de la tob buffer
    TOBBuffer := TOB.Create('', nil, -1);
//
    NomsTable := '';
    NomsFichier := '';
    for i_ind1 := 0 to TobTable.NbChamps - 1 do
        NomsTable := NomsTable + TobTable.GetNomChamp(i_ind1) + ';';
    for i_ind1 := 0 to StringGridFichier.RowCount - 1 do
        if StringGridFichier.Cells[0, i_ind1] <> '' then
            NomsFichier := NomsFichier + StringGridFichier.Cells[0, i_ind1] + ';';
//  ouverture du fichier de données
    Nb_Lus := 0;
    Nb_Ecrit := 0;
    ClePrec := '';
    st_trav1 := StringGrid1.Cells[i_indfic, 0] + '.pgi';
    AssignFile (Fichier, st_trav1);
    Reset (Fichier);  // equivalent open file
//  calcul nb enreg du fichier de données
    readln (Fichier,st_enreg);
//  Modif 07/02/02 rapprochement GC - GRC fichier vide
    if trim (st_enreg ) = '' then
    begin
    CloseFile (Fichier);
    exit;
    end
    else
//  Fin Modif 07/02/02 rapprochement GC - GRC
    begin
    FindFirst(st_trav1, faAnyFile, SearchRec);
    if NbMax <> High(i_ind1) then
        i_sizefic := Min(NbMax, Trunc(SearchRec.Size / Length(st_enreg)))
    else
        i_sizefic := Trunc(SearchRec.Size / Length(st_enreg));
    end;
    InitMove (i_sizefic, '');
    Reset (Fichier);  // equivalent open file
    ListeFichiersSeq.ItemIndex := i_indfic;
    ListeFichiersSeq.Selected[i_indfic] := True;
//    TOBALire := TOB.Create('', nil, -1);
    for i_ind1 := 0 to 999 do NumChArray[i_ind1] := 0;
    for i_ind1 := 0 to 999 do CleArray[i_ind1] := '';
    for i_ind1 := 0 to 999 do TOBArray[i_ind1] := nil;

//
//   ATTENTION : SPECIFIQUE POUR SQL 7
//
    if RGSortie.ItemIndex <> 0 then
        begin
        PrepareListe(RGSortie.ItemIndex, TobTable, SQL_Champs);
        if SQL_Champs <> '' then
            begin
            writeln(FicSeq, SQL_Champs);
            flush(FicSeq);
            end;
        end;
//  initialisation de la Tob
//    TobTable.InitValeurs;
    Existe_Tablette := False;
    for i_ind1 := 0 to StringGridAsso.RowCount - 1 do
        begin
        NumChArray[i_ind1] := TobTable.GetNumChamp(StringGridAsso.Cells[0, i_ind1]);
        if Copy(StringGridAsso.Cells[0, i_ind1], 0, 8) = 'Tablette' then Existe_Tablette := True;
        end;
//  traitement d'un enreg

    BeginTrans;

    while (not EOF(Fichier)) and (Nb_Ecrit <= NbMax) do
    Begin
        readln (Fichier,st_enreg);     // lecture des données
//        writeln (FicLog, 'Lecture enreg à ' + DateTimeToStr(Now));
        Label3.Caption := IntToStr(nb_lus);
        SQL_Insert := 'Insert into ' + Repertoire.Text;
        SQL_Update := 'Update ' + Repertoire.Text + ' set ';
        SQL_Champs := '';
        SQL_Datas  := '';
        SQL_Sep    := '';
        if (st_MAJ = 'CC') or (st_MAJ = 'MC') then
            begin
//  initialisation de la Tob
            TobTable.InitValeurs;
            for i_ind1 := 0 to TobTable.NbChamps - 1 do
                begin
                st_trav2 := TobTable.GetNomChamp(i_ind1);
                st_trav1 := Copy(st_trav2, Pos('_',st_trav2) + 1, 255);
                if st_trav1 = 'SOCIETE' then TobTable.PutValue(st_trav2, V_PGI.CodeSociete);
                if st_trav1 = 'ETABLISSEMENT' then TobTable.PutValue(st_trav2, VH^.EtablisDefaut);
                if st_trav1 = 'UTILISATEUR' then TobTable.PutValue(st_trav2, V_PGI.User);
                end;
            end
        else
//  traitement et recuperation des donnees de cle de la table pour les mises a jour
        begin
            st_cletable := '';
//  Modif 07/02/02 rapprochement GC - GRC  clé sur plusieurs champ table
            st_trav3 := Edit1.Text ;
            st_trav1:=uppercase(ReadTokenSt(st_trav3)) ;
            NomChampTable := Copy(st_trav1 , 0, Pos('=', st_trav1)-1);
            st_trav1 := Trim(Copy(st_trav1, Pos('=', st_trav1) + 1, 255));
//  fin Modif 07/02/02 rapprochement GC - GRC  clé sur plusieurs champ table
            while st_trav1 <> '' do
            begin
                if Pos('+', st_trav1) <> 0 then
                    NomChampSeq := Copy(st_trav1, 0, Pos('+', st_trav1) - 1)
                else
                    NomChampSeq := st_trav1;
                if (Copy(NomChampSeq, 0, 3) = st_racine) or (NomChampSeq[1] = '"') then
                begin
                    for i_ind2 := 0 to StringGridFichier.RowCount - 1 do
                        if NomChampSeq = StringGridFichier.Cells[0, i_ind2] then
                        begin
                            st_trav2 := StringGridFichier.Cells[2, i_ind2];
                            break;
                        end;
                    Extraire_Champ(1, '', st_trav2, ValeurSeq, ChampSeq);
                    Traiter_Champ_NEGOCE('', NomChampSeq, ChampSeq);
                    Traiter_Champ_PGI(NomChampTable, NomChampSeq, Variant(ChampSeq));
                    if Pos(NomChampSeq, Edit1.Text) <> 0 then
                        st_cletable := st_cletable + VarAsType(ChampSeq, varString);
                    if st_cletable <> ClePrec then TobTable.LoadDB(FALSE);
                end;

                if Pos('+', st_trav1) <> 0 then
                      st_trav1 := Copy(st_trav1, Pos('+', st_trav1) + 1, 255)
//  Modif 07/02/02 rapprochement GC - GRC  clé sur plusieurs champ table
                else if trim(st_trav3) <> '' then
                begin
                      TobTable.PutValue(NomChampTable, Variant(ChampSeq));
                      st_trav1:=uppercase(ReadTokenSt(st_trav3));
                      NomChampTable := Copy(st_trav1 , 0, Pos('=', st_trav1)-1);
                      st_trav1 := Trim(Copy(st_trav1, Pos('=', st_trav1) + 1, 255));
                      st_cletable := '';
                end
//  Fin Modif 07/02/02 rapprochement GC - GRC
                else  st_trav1 := '';

            end;
//            st_trav1 := Copy(Edit1.Text, 0, Pos('=', Edit1.Text) - 1);
//            st_trav2 := 'SELECT * FROM ' + Repertoire.Text + ' WHERE ' + st_trav1 +
//                        ' = "' + st_cletable + '"';
//            Query1 := OpenSql(st_trav2, False);
            TobTable.PutValue(NomChampTable, Variant(st_cletable));
            TobTable.LoadDB(FALSE);
//            Ferme(Query1);

        end;
//  Modif 06/02/02 rapprochement GC - GRC
// $ENDIF
//  Fin Modif 06/02/02 rapprochement GC - GRC
        l_Suivant := False;
//  traitement des champs en correspondance directe avec la table
        for i_ind1 := 0 to ListeAssoCh.Items.Count - 1 do
        begin
            Application.ProcessMessages;
            if l_Arret then Break;
            NomChampTable := StringGridAsso.Cells[0, i_ind1];
            if Copy(NomChampTable, 0, 8) <> 'Tablette' then
            begin
                i_ind2 := 1;
                while StringGridAsso.Cells[i_ind2, i_ind1] <> '' do
                begin
                    NomChampSeq := StringGridAssoCh.Cells[i_ind2, i_ind1];
                    if Copy(NomChampSeq, 2, 4) = 'calc' then
                    begin
                        ValeurSeq := NomChampSeq;
                        Break;
                    end;
                    if Pos('{SI', NomChampSeq) <> 0 then
                    begin
                        ValeurSeq := NomChampSeq;
                        Break;
                    end;
                    if (Copy(NomChampSeq, 0, 3) = st_racine) or (NomChampSeq[1] = '"') then
                    begin
                        if Copy(NomChampSeq, 4, 5) = 'PARAM' then
                            st_trav1 := '"' + StringGridAssoStruct.Cells[i_ind2, i_ind1] + '"'
                            else
                            st_trav1 := StringGridAssoStruct.Cells[i_ind2, i_ind1];
                        Extraire_Champ(i_ind2, NomChampTable, st_trav1, ValeurSeq, ChampSeq);
                        Traiter_Champ_NEGOCE(NomChampTable, NomChampSeq, ChampSeq);
                        if VarType(ChampSeq) = varDate then
                            ValeurSeq := ChampSeq
                        else
                            ValeurSeq := ValeurSeq + ChampSeq;
                    end;
                    i_ind2 := i_ind2 + 1;
                end;

                if VarType(ValeurSeq) = varString then
                begin
                    if Trim(ValeurSeq) = '' then Continue;
                    if Copy(ValeurSeq, 2, 4) = 'calc' then
                        Traiter_Calcul(NomChampTable, ValeurSeq)
                    else if Pos('{SI', ValeurSeq) <> 0 then
                        ValeurSeq := Evaluer_Si(ValeurSeq, NomChampTable)
                    else
                        ValeurSeq := TrimRight(ValeurSeq);
                end;
//  ValeurSeq = SUIVANT donc on est sur un SI dont le résultat est SUIVANT. On abandonne
//  l'enregistrement en cours
                if Pos('SUIVANT', ValeurSeq) <> 0 then
                    begin
                    l_Suivant := True;
                    Break;
                    end;
                TypeChamp := StringGridDECHAMPS.Cells[1, i_ind1];
//{$IFNDEF GRC}  Modif 07/02/02 rapprochement GC - GRC
                Champ_Tablette(NomChampTable, ValeurSeq, i_ind1);
//{$ENDIF}  Fin Modif 07/02/02 rapprochement GC - GRC
                if (NomChampSeq[1] = '"') and (Copy(NomChampSeq, 2, 4) <> 'calc') then
                begin
//  Modif 06/02/02 rapprochement GC - GRC
//                    if TypeChamp = 'DOUBLE' then
                    if (TypeChamp = 'DOUBLE') or (TypeChamp = 'RATE') then
//  Fin Modif 06/02/02 rapprochement GC - GRC
                    begin
                        i_PInt := 0;
                        i_PDec := 0;
                        if Pos('.', ValeurSeq) <> 0 then
                        begin
                            i_PInt := StrToInt(Copy(ValeurSeq, 0, Pos('.', ValeurSeq) - 1));
                            i_PDec := StrToInt(Copy(ValeurSeq, Pos('.', ValeurSeq) + 1, 255));
                            i_decimale := Length(ValeurSeq) - Pos('.', ValeurSeq);
                        end
                        else
                        begin
                            i_PInt := StrToInt(ValeurSeq);
                            i_decimale := 0;
                        end;
                        i_Sens := 1;
                        if i_PInt < 0 then i_Sens := -1;
                        case i_decimale of
                        0 : ChampSeq := i_PInt;
                        1 : ChampSeq := i_PInt + ((i_PDec / 10) * i_Sens);
                        2 : ChampSeq := i_PInt + ((i_PDec / 100) * i_Sens);
                        3 : ChampSeq := i_PInt + ((i_PDec / 1000) * i_Sens);
                        4 : ChampSeq := i_PInt + ((i_PDec / 10000) * i_Sens);
                        5 : ChampSeq := i_PInt + ((i_PDec / 100000) * i_Sens);
                        6 : ChampSeq := i_PInt + ((i_PDec / 1000000) * i_Sens);
                        7 : ChampSeq := i_PInt + ((i_PDec / 10000000) * i_Sens);
                        8 : ChampSeq := i_PInt + ((i_PDec / 100000000) * i_Sens);
                        9 : ChampSeq := i_PInt + ((i_PDec / 1000000000) * i_Sens);
                        end;
                    end;
                    if TypeChamp = 'INTEGER' then ValeurSeq := StrToInt(ValeurSeq);
                    if TypeChamp = 'DATE' then ValeurSeq := StrToDate(ValeurSeq);
                end;
                if TypeChamp = 'BLOB' then
                begin
                    ChampSeq := TobTable.GetValue(NomChampTable);
                    Memo1.Lines.Text := VarAsType(ChampSeq, varString);
//  Modif 07/02/02 rapprochement GC - GRC
                    st_cletable := TobTable.Cle1;
//  Modif fin 07/02/02 rapprochement GC - GRC
                    if st_cletable = ClePrec then
                        Memo1.Lines.Append(VarAsType(ValeurSeq,  varString))
                    else
                    begin
                        Memo1.Lines.Text := VarAsType(ValeurSeq,  varString);
                        ClePrec := st_cletable;
                    end;
                    ValeurSeq := Memo1.Lines.Text;
                end;
                Ecrire_Tob(NumChArray[i_ind1], ValeurSeq);
                Traiter_Champ_PGI(NomChampTable, NomChampSeq, ValeurSeq);
            end;
        end;

        if l_Arret then Break;
        if l_Suivant then
            begin
            Nb_Lus := Nb_Lus + 1;
            Continue;
            end;
//  traitement des champs definis en tablette
        if Existe_Tablette then
        begin
            for i_ind1 := 0 to ListeAssoCh.Items.Count - 1 do
            begin
                Application.ProcessMessages;
                if l_Arret then Break;
                NomChampTable := StringGridAsso.Cells[0, i_ind1];
                if Copy(NomChampTable, 0, 8) = 'Tablette' then
                begin
                    st_trav1 := StringGridAssoStruct.Cells[2, i_ind1];
                    TabAbrege := '';
                    if st_trav1 <> '' then
                    begin
                        Extraire_Champ(2, NomChampTable, st_trav1, ValeurSeq, ChampSeq);
                        if VarType(ChampSeq) = varString then
                            TabAbrege := TrimRight(ChampSeq)
                    end;
                    st_trav1 := StringGridAssoStruct.Cells[1, i_ind1];
                    Extraire_Champ(1, NomChampTable, st_trav1, ValeurSeq, ChampSeq);
                    if VarType(ChampSeq) = varString then
                        TabLibelle := TrimRight(ChampSeq)
                    else
                        TabLibelle := '';
    //  si le champ associe est bien dans le fichier en cours de traitement...
                    if Copy(ValeurSeq, 0, 3) = st_racine then
                        Traiter_Tablette(ValeurSeq, TabLibelle, TabAbrege);
                end;
            end;
        end;
        if l_Arret then Break;
//        if (TomTable.VerifTOB( TobTable )) then  //PCS
            ImportEnreg (TobTable, st_MAJ);
        inc(Nb_Ecrit);
        Nb_Lus := Nb_Lus + 1;
        MoveCur (False);
    end;
    Try
        CommitTrans;
    Except
        RollBack ;
    end;
    FiniMove ();
    CloseFile (Fichier);
    if l_Arret then
        begin
        writeln (FicLog, 'Traitement arreté par l''utilisateur');
        flush (FicLog);
        end;
    writeln (FicLog, IntToStr(Nb_Lus - 1) + ' lus, ' + IntToStr(Nb_Ecrit - 1) + ' écrits');
    flush (FicLog);
    i_LusTotal := i_LusTotal + Nb_Lus;
    TobBuffer.Free;
end;

//----------------------------------------------------------------------------
procedure TFRecupNeg.Ecrire_Tob(NumChamp : integer; ValeurSeq : Variant);
begin
try
    Case VarType(TobTable.GetValeur(NumChamp)) of
    varSmallint : TobTable.PutValeur(NumChamp, VarAsType(ValeurSeq, varSmallInt));
    varInteger  : TobTable.PutValeur(NumChamp, VarAsType(ValeurSeq, varInteger));
    varSingle   : TobTable.PutValeur(NumChamp, VarAsType(ValeurSeq, varSingle));
    varDouble   : TobTable.PutValeur(NumChamp, VarAsType(ValeurSeq, varDouble));
    varDate     : TobTable.PutValeur(NumChamp, VarAsType(ValeurSeq, varDate));
    varString   : TobTable.PutValeur(NumChamp, VarAsType(ValeurSeq, varString));
    else
        TobTable.PutValeur(NumChamp, ValeurSeq);
    end;
//    TobTable.PutValue(NomChampTable, ValeurSeq);
except
    on EVariantError do GereConvertError(TobTable.GetNomChamp(NumChamp), ValeurSeq);
    on EConvertError do GereConvertError(TobTable.GetNomChamp(NumChamp), ValeurSeq);
end;
end;

//----------------------------------------------------------------------------
procedure TFRecupNeg.Extraire_Champ(i_ind1 : Integer; NomChampTable, st_trav1 : string; var ValeurSeq : Variant;
                                    var ChampSeq : Variant);
var
    st_format, st_chassoc, st_trav2 : string;
    i_offset, i_longueur, i_decimale : integer;
    i_PInt, i_PDec, i_Sens : integer;

begin

//  si le 1er parametre est egal a 0, on doit chercher les infos du champ avant de continuer.
//  le nom du champ est dans NomChampTable.
Label4.Caption := NomChampTable;
if i_ind1 = 0 then
    begin
    if NomChampTable[1] = '"' then
        st_trav1 := Copy(NomChampTable, 2, Length(NomChampTable) - 2)
        else
        for i_offset := 0 to StringGridFichier.RowCount - 1 do
            if NomChampTable = StringGridFichier.Cells[0, i_offset] then
            begin
                if Pos('PARAM', NomChampTable) <> 0 then
                   st_trav1 := '"' + StringGridFichier.Cells[2, i_offset] + '"'
                   else
                   st_trav1 := StringGridFichier.Cells[2, i_offset];
                break;
            end;
    NomChampTable := '';
    end;
//  si c'est une chaine on la recupere sans les guillemets
    if st_trav1[1] = '"' then
    begin
        if i_ind1 = 1 then ValeurSeq := '';
        ChampSeq := Copy(st_trav1, 2, Length(st_trav1) - 2);
    end
    else
    begin
//  recupere les differentes infos necessaires
        i_offset := StrToInt(ReadTokenSt(st_trav1));
        i_longueur := StrToInt(ReadTokenSt(st_trav1));
        i_decimale := StrToInt(ReadTokenSt(st_trav1));
        st_format := ReadTokenSt(st_trav1);
        st_chassoc := ReadTokenSt(st_trav1);

        if st_format = 'A' then
        begin
            if i_ind1 = 1 then ValeurSeq := '';
            ChampSeq := TrimRight(Copy(st_enreg, i_offset, i_longueur));
            if Copy(NomChampTable, 0, 8) = 'Tablette' then ValeurSeq := st_chassoc;
            Exit;
        end;
        if st_format = 'D' then
        begin
            if i_ind1 = 1 then ValeurSeq := '';
            st_trav2 := TrimRight(Copy(st_enreg, i_offset, i_longueur));
            ChampSeq := '';
            while Pos('/', st_trav2) <> 0 do
            begin
                ChampSeq := ChampSeq + Copy(st_trav2, 0, Pos('/', st_trav2) - 1);
                st_trav2 := Copy(st_trav2, Pos('/', st_trav2) + 1, 255);
            end;
            ChampSeq := ChampSeq + st_trav2;
            if not IsNumeric(ChampSeq) then
                begin
                GereConvertError(NomChampTable, ChampSeq);
                Exit
                end;
            while Pos(' ',ChampSeq) <> 0 do
                ChampSeq := Copy(ChampSeq, 0, Pos(' ',ChampSeq) - 1) +
                            Copy(ChampSeq, Pos(' ',ChampSeq) + 1, 255);
            if StrToInt(ChampSeq) = 0 then
//  Date a 0 : on met le 31/12/2099 par defaut
                ChampSeq := Str8ToDate('31122099', True)
            else
                if Length(ChampSeq) = 6 then
                    try
                        ChampSeq := Str6ToDate(ChampSeq, 20);
                    except
                      on EVariantError do GereConvertError(NomChampTable, ChampSeq);
                      on EConvertError do GereConvertError(NomChampTable, ChampSeq);
                    end
                else
                    try
                        ChampSeq := Str8ToDate(ChampSeq, True);
                    except
                      on EVariantError do GereConvertError(NomChampTable, ChampSeq);
                      on EConvertError do GereConvertError(NomChampTable, ChampSeq);
                    end;
            Exit;
        end;
        if st_format = 'N' then
        begin
            if i_ind1 = 1 then ValeurSeq := 0.0;
            i_PDec := 0;
            ChampSeq := 0;
            st_trav1 := TrimLeft(Copy(st_enreg, i_offset, i_longueur));
            i_Sens := 1;
            if st_trav1 = '' then Exit;
            if st_trav1[1] = '-' then i_Sens := -1;
            if IsNumeric(st_trav1) then
                begin
                while Pos(',', st_trav1) <> 0 do st_trav1[Pos(',', st_trav1)] := '.';
                if Pos('.', st_trav1) <> 0 then
                begin
                    if Trim(Copy(st_trav1, 0, Pos('.', st_trav1) - 1)) <> '' then
                        i_PInt := StrToInt(Copy(st_trav1, 0, Pos('.', st_trav1) - 1))
                    else
                    begin
                        GereConvertError(NomChampTable, st_trav1);
                        Exit;
                    end;
                    if Trim(Copy(st_trav1, Pos('.', st_trav1) + 1, 255)) <> '' then
                        i_PDec := StrToInt(Copy(st_trav1, Pos('.', st_trav1) + 1, 255))
                    else
                    begin
                        GereConvertError(NomChampTable, st_trav1);
                        Exit;
                    end;
                end
                else
                    try
                    i_PInt := StrToInt(st_trav1);
                    except
                      on EVariantError do GereConvertError(NomChampTable, ChampSeq);
                      on EConvertError do GereConvertError(NomChampTable, ChampSeq);
                    end;
                end
                else
                begin
                GereConvertError(NomChampTable, st_trav1);
                Exit;
                end;
            case i_decimale of
            0 : ChampSeq := i_PInt;
            1 : ChampSeq := i_PInt + ((i_PDec / 10) * i_Sens);
            2 : ChampSeq := i_PInt + ((i_PDec / 100) * i_Sens);
            3 : ChampSeq := i_PInt + ((i_PDec / 1000) * i_Sens);
            4 : ChampSeq := i_PInt + ((i_PDec / 10000) * i_Sens);
            5 : ChampSeq := i_PInt + ((i_PDec / 100000) * i_Sens);
            6 : ChampSeq := i_PInt + ((i_PDec / 1000000) * i_Sens);
            7 : ChampSeq := i_PInt + ((i_PDec / 10000000) * i_Sens);
            8 : ChampSeq := i_PInt + ((i_PDec / 100000000) * i_Sens);
            9 : ChampSeq := i_PInt + ((i_PDec / 1000000000) * i_Sens);
            end;
            if stArrondi = 'O' then
                ChampSeq := Arrondi(ChampSeq, i_decimale);
            Exit;
        end;
    end;
end;

//----------------------------------------------------------------------------
procedure TFRecupNeg.Traiter_Tablette(ChampAssocie, TabLibelle, TabAbrege : string);
var
    st_trav1, st_trav2, st_Champ : string;
    i_ind1, i_offset : integer;
    l_Trouve : boolean;
    v_trav1, v_cle : variant;
    Tob_ValCombo, TobTemp : Tob;

begin
//  recherche du champ table correspondant
    l_Trouve := False;
    if ChampAssocie <> '' then
    begin
        for i_ind1 := 0 to StringGridAssoCh.RowCount - 1 do
        begin
            if StringGridAssoCh.Cells[0, i_ind1] = '' then Break;
            for i_offset := 1 to StringGridAssoCh.ColCount - 1 do
            begin
                if StringGridAssoCh.Cells[i_offset, i_ind1] = '' then Break;
                if StringGridAssoCh.Cells[i_offset, i_ind1] = ChampAssocie then
                begin
                    l_Trouve := True;
                    Break;
                end;
            end;
            if l_Trouve then Break;
        end;
    end;
    if not l_Trouve then Exit;
//  champ trouve : est ce un Combo ?
    st_Champ := StringGridAssoCh.Cells[0, i_ind1];
    v_cle := TobTable.GetValue(st_Champ);
//  Pas de données associées récupérées, on sort
    if v_cle = '' then Exit;
    TobTemp := Tob_DeChamps.FindFirst(['DH_NOMCHAMP'], [st_Champ], False);
    st_trav1 := TobTemp.GetValue('DH_TYPECHAMP');
//    if st_trav1 <> 'COMBO' then Exit;
//  c'est un combo, on renseigne la tablette si le type est CO
    st_trav1 := Copy(st_Champ, Pos('_', st_Champ) + 1, 255);
    TobTemp := Tob_DeCombos.FindFirst(['DO_NOMCHAMP'], [st_trav1], False);
    if TobTemp = nil then Exit;
    st_trav1 := TobTemp.GetValue('DO_PREFIXE');
    st_trav2 := TobTemp.GetValue('DO_WHERE');
    st_trav2 := Copy(st_trav2, Pos('"', st_trav2) + 1, 255);
    st_trav2 := Copy(st_trav2, 0, Pos('"', st_trav2) - 1);
    if st_trav1 = 'CC' then
    begin
        Tob_ValCombo := Tob_ChoixCod.FindFirst([st_trav1+'_TYPE', st_trav1+'_CODE'],
                                               [st_trav2, Copy(v_cle, 1, 3)], False);
        if Tob_ValCombo <> nil then Exit;
{        While Tob_ValCombo <> nil do
            begin
            st_trav3 := Tob_ValCombo.GetValue(st_trav1+'_CODE');
            if st_trav3 = Copy(v_cle, 1, 3) then Exit;
            Tob_ValCombo := Tob_ChoixCod.FindNext([st_trav1+'_TYPE'], [st_trav2], False);
            end;  }
        Tob_ValCombo := TOB.Create('CHOIXCOD',Tob_ChoixCod,-1);
        v_trav1 := st_trav2;
        Tob_ValCombo.PutValue(st_trav1 + '_TYPE', v_trav1);
        Tob_ValCombo.PutValue(st_trav1 + '_CODE', Copy(v_cle, 1, 3));
        v_trav1 := TabLibelle;
        Tob_ValCombo.PutValue(st_trav1 + '_LIBELLE', v_trav1);
        v_trav1 := TabAbrege;
        Tob_ValCombo.PutValue(st_trav1 + '_ABREGE', v_trav1);
    end
    else if st_trav1 = 'YX' then
    begin
        Tob_ValCombo := Tob_ChoixExt.FindFirst([st_trav1+'_TYPE', st_trav1+'_CODE'],
                                               [st_trav2, Copy(v_cle, 1, 6)], False);
        if Tob_ValCombo <> nil then Exit;
{        While Tob_ValCombo <> nil do
            begin
            if Tob_ValCombo.GetValue(st_trav1+'_CODE') = Copy(v_cle, 1, 6) then Exit;
            Tob_ValCombo := Tob_ChoixExt.FindNext([st_trav1+'_TYPE'], [st_trav2], False);
            end; }
        Tob_ValCombo := TOB.Create('CHOIXEXT',Tob_ChoixExt,-1);
        v_trav1 := st_trav2;
        Tob_ValCombo.PutValue(st_trav1 + '_TYPE', v_trav1);
        Tob_ValCombo.PutValue(st_trav1 + '_CODE', Copy(v_cle, 1, 6));
        v_trav1 := TabLibelle;
        Tob_ValCombo.PutValue(st_trav1 + '_LIBELLE', v_trav1);
        v_trav1 := TabAbrege;
        Tob_ValCombo.PutValue(st_trav1 + '_ABREGE', v_trav1);
    end
    else
    begin
        Tob_ValCombo := Tob_Commun.FindFirst(['CO_TYPE','CO_CODE'],
                                             [st_trav2, Copy(v_cle, 1, 3)], False);
//        if not PresenceComp('COMMUN', 'CO_TYPE,CO_CODE', st_trav2 + ',' + Copy(v_cle, 1, 3)) then
        if Tob_ValCombo = nil then
            if Pos(st_trav2 + ' - ' + v_cle, ErreurCommun.Text) = 0 then
                ErreurCommun.Add(st_trav2 + ' - ' + v_cle);
    end;
//    Tob_ValCombo.Free;
end;

{=============================================================================}
procedure TFRecupNeg.Champ_Tablette(NomChampTable, Valeur : string; i_ind1 : integer);
var
    st_trav1, st_trav2 : string;
    Tob_ValCombo : Tob;

begin
    st_trav1 := StringGridDECHAMPS.Cells[2, i_ind1];
    if (st_trav1 <> 'CC') and (st_trav1 <> 'CO') and (st_trav1 <> 'YX') then Exit;
    st_trav2 := StringGridDECHAMPS.Cells[3, i_ind1];
    st_trav2 := Copy(st_trav2, Pos('"', st_trav2) + 1, 255);
    st_trav2 := Copy(st_trav2, 0, Pos('"', st_trav2) - 1);
    if st_trav1 = 'CC' then
    begin
        Tob_ValCombo := Tob_ChoixCod.FindFirst([st_trav1+'_TYPE', st_trav1+'_CODE'],
                                               [st_trav2, Copy(Valeur, 1, 3)], False);
        if Tob_ValCombo <> nil then Exit;
//        While Tob_ValCombo.GetValue(st_trav1+'_TYPE') = st_trav2 do
{        While Tob_ValCombo <> nil do
            begin
            st_trav3 := Tob_ValCombo.GetValue(st_trav1+'_CODE');
            if st_trav3 = Copy(Valeur, 1, 3) then Exit;
            Tob_ValCombo := Tob_ChoixCod.FindNext([st_trav1+'_TYPE'], [st_trav2], False);
            end;  }
        Tob_ValCombo := TOB.Create('CHOIXCOD',Tob_ChoixCod,-1);
        Tob_ValCombo.PutValue(st_trav1 + '_TYPE', st_trav2);
        Tob_ValCombo.PutValue(st_trav1 + '_CODE', Copy(Valeur, 1, 3));
        Tob_ValCombo.PutValue(st_trav1 + '_LIBELLE', 'Libellé à renseigner');
        Tob_ValCombo.PutValue(st_trav1 + '_ABREGE', '?????');
//        Tob_ValCombo.InsertOrUpdateDB;
    end
    else if st_trav1 = 'YX' then
    begin
        Tob_ValCombo := Tob_ChoixExt.FindFirst([st_trav1+'_TYPE', st_trav1+'_CODE'],
                                               [st_trav2, Copy(Valeur, 1, 6)], False);
        if Tob_ValCombo <> nil then Exit;
//        While Tob_ValCombo.GetValue(st_trav1+'_TYPE') = st_trav2 do
{        While Tob_ValCombo <> nil do
            begin
            if Tob_ValCombo.GetValue(st_trav1+'_CODE') = Copy(Valeur, 1, 6) then Exit;
            Tob_ValCombo := Tob_ChoixExt.FindNext([st_trav1+'_TYPE'], [st_trav2], False);
            end;  }
        Tob_ValCombo := TOB.Create('CHOIXEXT',Tob_ChoixExt,-1);
        Tob_ValCombo.PutValue(st_trav1 + '_TYPE', st_trav2);
        Tob_ValCombo.PutValue(st_trav1 + '_CODE', Copy(Valeur, 1, 6));
        Tob_ValCombo.PutValue(st_trav1 + '_LIBELLE', 'Libellé à renseigner');
        Tob_ValCombo.PutValue(st_trav1 + '_ABREGE', '?????');
//        Tob_ValCombo.InsertOrUpdateDB;
    end
    else
    begin
//        Tob_ValCombo := TOB.Create('COMMUN',Nil,-1);
        Tob_ValCombo := Tob_Commun.FindFirst(['CO_TYPE', 'CO_CODE'],
                                             [st_trav2, Copy(valeur, 1, 3)], False);
//        if not PresenceComp('COMMUN', 'CO_TYPE,CO_CODE', st_trav2 + ',' + Copy(v_cle, 1, 3)) then
        if Tob_ValCombo = nil then
            if Pos(st_trav2 + ' - ' + Valeur, ErreurCommun.Text) = 0 then
                ErreurCommun.Add(st_trav2 + ' - ' + Valeur);
    end;
//    Tob_ValCombo.Free;
end;

{=============================================================================}
procedure TFRecupNeg.Recherche_Tablette(NomChampTable : string; i_ind1 : integer);
var
    st_trav1 : string;
    TobTemp : TOB;

begin
//  champ trouve : est ce un Combo ?
{    st_trav1 := 'SELECT * FROM DECHAMPS WHERE DH_NOMCHAMP = "' + NomChampTable + '"';
    Query1 := OpenSql(st_trav1, False);
    Tob_DeChamps.LoadDetailDB('DECHAMPS','','',Query1,FALSE);
    Tob_DeChamps.PutStringsDetail(StringGrid4.Cols[0],'DH_TYPECHAMP');
    Ferme(Query1); }
    TobTemp := Tob_DeChamps.FindFirst(['DH_NOMCHAMP'], [NomChampTable], False);
    if TobTemp = nil then Exit;
    StringGridDECHAMPS.Cells[1, i_ind1] := TobTemp.GetValue('DH_TYPECHAMP');
    if TobTemp.GetValue('DH_TYPECHAMP') <> 'COMBO' then Exit;
// temporaire en attendant de regler le pb des multiples tablettes devises
    if Copy(NomChampTable, Pos('_', NomChampTable) + 1, 255) = 'DEVISE' then Exit;
//  c'est un combo, on renseigne la tablette si le type est CO
{    st_trav1 := 'SELECT * FROM DECOMBOS WHERE DO_NOMCHAMP = "' +
                Copy(NomChampTable, Pos('_', NomChampTable) + 1, 255) + '"';
    Query1 := OpenSql(st_trav1, False);
    if not Query1.Eof then
        begin
        Tob_DeCombos.LoadDetailDB('DECOMBOS','','',Query1,FALSE);
        Tob_DeCombos.PutStringsDetail(StringGrid4.Cols[0], 'DO_PREFIXE');
        StringGridDECHAMPS.Cells[2, i_ind1] := StringGrid4.Cells[0, 0];
        Tob_DeCombos.PutStringsDetail(StringGrid4.Cols[0], 'DO_WHERE');
        StringGridDECHAMPS.Cells[3, i_ind1] := StringGrid4.Cells[0, 0];
        end;
    Ferme(Query1);  }
    st_trav1 := Copy(NomChampTable, Pos('_', NomChampTable) + 1, 255);
    TobTemp := Tob_DeCombos.FindFirst(['DO_NOMCHAMP'], [st_trav1], False);
    if TobTemp <> nil then
        begin
        StringGridDECHAMPS.Cells[2, i_ind1] := TobTemp.GetValue('DO_PREFIXE');
        StringGridDECHAMPS.Cells[3, i_ind1] := TobTemp.GetValue('DO_WHERE');
        end;
end;

{=============================================================================}
function TFRecupNeg.PresenceComp (Fichier,Champs,Valeurs : String ) : Boolean ;
Var
    Q : TQuery ;
    st_trav1 : string;

begin
    st_trav1 := 'Select ' + Champs + ' from ' + Fichier + ' Where ';
    while Champs <> '' do
        st_trav1 := st_trav1 + ReadTokenStV(Champs) + ' = "' + ReadTokenStV(Valeurs) + '" and ';
    st_trav1 := Copy(st_trav1, 0, Length(st_trav1) - 5);
    Q:=OpenSQL(st_trav1, TRUE,-1,'',true) ;
    Result:=(not Q.EOF) ;
    Ferme(Q);
    SourisNormale ;
end;

{=============================================================================}
procedure TFRecupNeg.Enreg_Log_Erreur;
var
    i_ind1 : integer;
    st_trav1 : string;

begin
    for i_ind1 := 0 to ErreurCommun.Count - 1 do
    begin
        st_trav1 := ErreurCommun.Strings[i_ind1];
        writeln (FicLog, 'AVERTISSEMENT : ' + st_trav1 + ' n''existe ' +
                         'pas dans COMMUN et est présent dans les données lues');
        flush (FicLog);
    end;
    ErreurCommun.Text := '';
end;

//----------------------------------------------------------------------------
procedure TFRecupNeg.Traiter_Champ_NEGOCE(NomChampTable : string; NomChampSeq : string;
                                          var ValeurSeq : Variant);
var
//    i_increment : integer;
    st_trav1, st_trav2 : string;

begin
//  Champ CODE ARTICLE
    if NomChampSeq = 'ART0010' then
    begin
        ValeurSeq := UpperCase(ValeurSeq);
        st_trav1 := ValeurSeq;
        st_trav2 := ValeurSeq;
{        While Presence ('ARTICLE', 'GA_ARTICLE', ValeurSeq) Do
        Begin
            inc (i_increment);
            ValeurSeq := ValeurSeq + Format ('%.3u', [i_increment]);
        End;  }
        if Incremente('ARTICLE', 'GA_ARTICLE', False, st_trav2) <> 0 then
        begin
            ValeurSeq := st_trav2;
            writeln (FicLog, 'Code article ' + st_trav1 + ' remplacé par ' + ValeurSeq);
            flush (FicLog);
            if Correspond.IndexOf('ARTICLE;' + st_trav1 + ';' + ValeurSeq) < 0 then
                begin
                writeln (Corresp, 'ARTICLE;' + st_trav1 + ';' + ValeurSeq);
                flush (FicLog);
                Correspond.Add('ARTICLE;' + st_trav1 + ';' + ValeurSeq);
                end;
        end;
        Exit;
    end
//  Champ COMMISSIONNABLE
    else if NomChampSeq = 'ART0070' then
    begin
        ValeurSeq := UpperCase(ValeurSeq);
        if (ValeurSeq <> 'O') and (ValeurSeq <> 'N') then ValeurSeq := 'N';
        Exit;
    end
//  Champ TENUESTOCK
    else if NomChampSeq = 'ART0140' then
    begin
        ValeurSeq := UpperCase(ValeurSeq);
        if (ValeurSeq <> 'O') and (ValeurSeq <> 'N') then ValeurSeq := 'N';
        Exit;
    end
//  Champ Code Client
{    if (NomChampSeq = 'CLI0010') or (NomChampSeq = 'BCL0010') or (NomChampSeq = 'CLI0230')then
    begin
        st_trav1 := ParamGenData.Strings[ParamGenCle.IndexOf('Client')];
        st_trav1 := Copy(st_trav1, Pos('=', st_trav1) + 1, 255);
        ValeurSeq := UpperCase(st_trav1 + ValeurSeq);
        Exit;
    end;
//  Champ Code Fournisseur
    if (NomChampSeq = 'FOU0010')  or (NomChampSeq = 'BFO0010') then
    begin
        st_trav1 := ParamGenData.Strings[ParamGenCle.IndexOf('Fournisseur')];
        st_trav1 := Copy(st_trav1, Pos('=', st_trav1) + 1, 255);
        ValeurSeq := UpperCase(st_trav1 + ValeurSeq);
        Exit;
    end; }
//  Code lot Negoce dans le fichier Stock
    else if NomChampSeq = 'STO0010' then
    begin
//  stockage temporaire pour le ca&s des BLC eventuels
        st_temp1 := ValeurSeq;
        Exit;
    end
//  Modif 06/02/02 rapprochement GC - GRC
    else if (copy(NomChampTable,1,13) = 'RPR_RPRLIBMUL') and
            (ParamGenData.Strings[ParamGenCle.IndexOf('ModeCreation')]<>'C') then
    begin
       st_trav2 := TobTable.GetValue(NomChampTable);
       if (st_trav2 = '') or (st_trav2 = ';') then  ValeurSeq := ValeurSeq+';'
       else if (pos(ValeurSeq,st_trav2) = 0) then
       begin
          if (st_trav2[Length(st_trav2)] = ';') then ValeurSeq := st_trav2+ValeurSeq+';'
          else  ValeurSeq := st_trav2 + ';' + ValeurSeq + ';';
       end;
//  Fin Modif 06/02/02 rapprochement GC - GRC
    end;
end;

//----------------------------------------------------------------------------
procedure TFRecupNeg.Traiter_Champ_PGI(NomChampTable : string; NomChampSeq : string;
                                       var ValeurSeq : Variant);
var
    st_trav2, st_trav3, st_trav4 : string;
    pc_trav4 : string;
    NumChamp : integer;

begin
//  Code article
    NumChamp := TobTable.GetNumChamp(NomChampTable);

    st_trav2 := Copy(NomChampTable, 0, Pos('_', NomChampTable) - 1);
    st_trav3 := Copy(NomChampTable, Pos('_', NomChampTable) + 1, 255);
    if st_trav3 = 'AUXILIAIRE' then
    begin
        while Length(ValeurSeq) < LgCptAux do ValeurSeq := ValeurSeq + CarBourrage;
        Ecrire_Tob(NumChamp, ValeurSeq);
    end
    else if st_trav3 = 'NOMENCLATURE' then
    begin
        if Trim(ValeurSeq) = '' then Exit;
        pc_trav4 := Uppercase(CodeArticleUnique2(TobTable.GetValue(st_trav2 + '_' + st_trav3), ''));
        st_trav4 := pc_trav4;
        Ecrire_Tob(TobTable.GetNumChamp(st_trav2 + '_NOMENCLATURE'), pc_trav4);
        Exit;
    end
    else if st_trav3 = 'CODEARTICLE' then
    begin
        if Trim(ValeurSeq) = '' then Exit;
        pc_trav4 := Uppercase(CodeArticleUnique2(TobTable.GetValue(st_trav2 + '_' + st_trav3), ''));
        st_trav4 := pc_trav4;
        Ecrire_Tob(TobTable.GetNumChamp(st_trav2 + '_ARTICLE'), pc_trav4);
        Ecrire_Tob(TobTable.GetNumChamp(st_trav2 + '_CODEARTICLE'), UpperCase(ValeurSeq));
        if st_trav2 = 'GA' then Ecrire_Tob(TobTable.GetNumChamp('GA_STATUTART'),'UNI');
        Exit;
    end
    else if st_trav3 = 'ARTICLE' then
    begin
        if Trim(ValeurSeq) = '' then Exit;
//        pc_trav4 := Uppercase(CodeArticleUnique2(TobTable.GetValue(st_trav2 + '_' + st_trav3), ''));
        ValeurSeq := Uppercase(CodeArticleUnique2(ValeurSeq, ''));
        st_trav4 := ValeurSeq;
        Ecrire_Tob(TobTable.GetNumChamp(st_trav2 + '_ARTICLE'), ValeurSeq);
        if TobTable.FieldExists(st_trav2 + '_CODEARTICLE') then
        begin
            st_trav3 := TobTable.GetValue(st_trav2 + '_' + st_trav3);
            st_trav4 := '';
            pc_trav4 := Uppercase(CodeArticleGenerique2(st_trav3, st_trav4));
            st_trav4 := pc_trav4;
            Ecrire_Tob(TobTable.GetNumChamp(st_trav2 + '_CODEARTICLE'), pc_trav4);
            if st_trav2 = 'GA' then Ecrire_Tob(TobTable.GetNumChamp('GA_STATUTART'),'UNI');
        end;
        Exit;
    end
//  Code article du CATALOGUE
    else if NomChampTable = 'GCA_ARTICLE' then
    begin
        if Trim(TobTable.GetValue('GCA_REFERENCE')) = '' then
            Ecrire_Tob(TobTable.GetNumChamp('GCA_REFERENCE'),TobTable.GetValue('GCA_ARTICLE'));
        Exit;
    end
//  Champ DIMMASQUE
    else if NomChampSeq = 'ART0500' then
    begin
        Ecrire_Tob(NumChamp,'');
        Ecrire_Tob(TobTable.GetNumChamp('GA_DIMMASQUE'),ValeurSeq);
        Ecrire_Tob(TobTable.GetNumChamp('GA_STATUTART'),'GEN');
    end
    else if NomChampTable = 'T_LIBELLE' then
    begin
        Ecrire_Tob(TobTable.GetNumChamp('T_PHONETIQUE'),phoneme(TobTable.GetValue('T_LIBELLE')));
    end
    else if NomChampTable = 'T_FACTURE' then
    begin
        while Length(ValeurSeq) < LgCptAux do ValeurSeq := ValeurSeq + CarBourrage;
        Ecrire_Tob(NumChamp, ValeurSeq);
    end
    else if NomChampTable = 'T_PAYEUR' then
    begin
        while Length(ValeurSeq) < LgCptAux do ValeurSeq := ValeurSeq + CarBourrage;
        Ecrire_Tob(NumChamp, ValeurSeq);
    end;
end;

procedure TFRecupNeg.Controle_Ordre_De_Traitement;
begin
{    st_trav1 := ParamGenData.Strings[ParamGenCle.IndexOf('TypeDimVal')];
    BLC_Val := st_trav1;
    st_trav1 := Copy(st_trav1, Length(st_trav1), 1); }
//    st_trav1 := Copy(st_trav1, Pos('=', st_trav1) + 1, 255);

    if (ListeATraiter.Items.IndexOf('LIGNE') >= 0) and
       (ListeATraiter.Items.IndexOf('PIECE') >= 0) then
    begin
        if ListeATraiter.Items.IndexOf('PIECE') > ListeATraiter.Items.IndexOf('LIGNE') then
        begin
            ListeATraiter.Items.Exchange(ListeATraiter.Items.IndexOf('PIECE'),
                                         ListeATraiter.Items.IndexOf('LIGNE'));
        end;
    end;

{    if (ListeATraiter.Items.IndexOf('ARTICLE') >= 0) and (BLC_Val <> '') and
       (ListeATraiter.Items.IndexOf('BLC') = -1) then
    begin
//        ListeATraiter.Items.Insert(ListeATraiter.Items.IndexOf('ARTICLE'),'BLC');
        st_trav3 := 'CC_TYPE,CC_CODE';
        TobDim := TOB.Create('CHOIXCOD',Nil,-1);
        TobDim.PutValue('CC_TYPE', 'DIM');
        for i_ind1 := 1 to 5 do
        begin
            TobDim.PutValue('CC_CODE', 'DI' + IntToStr(i_ind1));
            st_trav2 := 'DIM,DI' + IntToStr(i_ind1);
            TobDim.PutValue('CC_LIBELLE', 'Dimension ' + IntToStr(i_ind1));
            TobDim.PutValue('CC_ABREGE', 'Dimension ' + IntToStr(i_ind1));
            if not PresenceComp('CHOIXCOD', st_trav3, st_trav2) then TobDim.InsertDB(Nil);
        end;
//        TobDim.Free;
//        TobDim := TOB.Create('CHOIXCOD',Nil,-1);
        TobDim.PutValue('CC_TYPE', 'GG' + st_trav1);
        TobDim.PutValue('CC_CODE', 'B');
        TobDim.PutValue('CC_LIBELLE', 'Bains');
        TobDim.PutValue('CC_ABREGE', 'Bains');
        TobDim.InsertOrUpdateDB;
        TobDim.PutValue('CC_TYPE', 'GG' + st_trav1);
        TobDim.PutValue('CC_CODE', 'L');
        TobDim.PutValue('CC_LIBELLE', 'Lots');
        TobDim.PutValue('CC_ABREGE', 'Lots');
        TobDim.InsertOrUpdateDB;
        TobDim.PutValue('CC_TYPE', 'GG' + st_trav1);
        TobDim.PutValue('CC_CODE', 'C');
        TobDim.PutValue('CC_LIBELLE', 'Coloris');
        TobDim.PutValue('CC_ABREGE', 'Coloris');
        TobDim.InsertOrUpdateDB;
        TobDim.Free;
        TobDim := TOB.Create('DIMMASQUE',Nil,-1);
        TobDim.PutValue('GDM_MASQUE', 'B');
        TobDim.PutValue('GDM_LIBELLE', 'Bains');
        TobDim.PutValue('GDM_TYPE' + st_trav1, 'B');
        TobDim.PutValue('GDM_POSITION' + st_trav1, 'CO1');
        TobDim.InsertOrUpdateDB;
        TobDim.PutValue('GDM_MASQUE', 'L');
        TobDim.PutValue('GDM_LIBELLE', 'Lots');
        TobDim.PutValue('GDM_TYPE' + st_trav1, 'L');
        TobDim.PutValue('GDM_POSITION' + st_trav1, 'CO1');
        TobDim.InsertOrUpdateDB;
        TobDim.PutValue('GDM_MASQUE', 'C');
        TobDim.PutValue('GDM_LIBELLE', 'Coloris');
        TobDim.PutValue('GDM_TYPE' + st_trav1, 'C');
        TobDim.PutValue('GDM_POSITION' + st_trav1, 'CO1');
        TobDim.InsertOrUpdateDB;
        TobDim.Free;
    end; }
end;

//------------------------------------------------------------------------------
//  Recherche dans la table NOMTABLE la(es) clé(s) NOMCHAMP de valeur(s) VALEURSEQ.
//  Si on trouve, on increment un compteur et on verifie l'existence du champ
//  NOMCHAMP avec la nouvelle VALEURSEQ. Le parametre REMPLACE permet de dire
//  si l'increment Remplace la valeur du dernier champ passé ou est ajouté à
//  Celui ci.
function TFRecupNeg.Incremente(NomTable, NomChamp : string; Remplace : Boolean;
                               var ValeurSeq : string) : Integer;
var
    i_increment : integer;
    st_trav1, st_trav2 : string;

begin
    i_increment := 0;
    ValeurSeq := UpperCase(ValeurSeq);
    st_trav1 := ValeurSeq;
    st_trav2 := '';
    if Remplace then
        while Pos(',', st_trav1) <> 0 do
            if st_trav2 = '' then
                st_trav2 := ReadTokenStV(st_trav1)
            else
                st_trav2 := st_trav2 + ',' + ReadTokenStV(st_trav1);
    st_trav1 := ValeurSeq;
    While PresenceComp (NomTable, NomChamp, ValeurSeq) Do
    Begin
        inc (i_increment);
        if Remplace then
            ValeurSeq := st_trav2 + ',' + Format ('%.3u', [i_increment])
        else
            ValeurSeq := st_trav1 + Format ('%.3u', [i_increment]);
    End;
    Result := i_increment;
end;

procedure TFRecupNeg.Traiter_Calcul(NomChampTable : string; var ValeurSeq : Variant);
var
    st_variables : array [0..50] of string;
    st_types : array [0..50] of char;
    i_ind1, i_variables, i_chaine : integer;
    st_trav1, st_trav2, st_trav3 : string;
    ch_trav1 : char;
    va_temp1, va_temp2 : variant;

//  Fonction determinant si la chaine est uniquement faite de chiffres et d'un point ou pas
    function Determine_Type : char;
    var
        i_ind1, i_offset : integer;

    begin
        Result := 'A';
        for i_ind1 := 66 to 92 do
            if Pos(Chr(i_ind1), st_trav2) <> 0 then Exit;
        for i_ind1 := 98 to 124 do
            if Pos(Chr(i_ind1), st_trav2) <> 0 then Exit;
        Result := 'N';
        i_offset := 0;
        for i_ind1 := 0 to Length(st_trav2) do
        begin
            if Copy(st_trav2, i_ind1, 1) = '.' then
                inc(i_offset);
        end;
        if i_offset > 1 then
        begin
            Result := 'A';
        end
        else
        begin
            if i_offset = 0 then
                st_trav3 := st_trav2 + DecimalSeparator + '0'
                else
                begin
                st_trav3 := Copy(st_trav2, 0, Pos('.', st_trav2) - 1);
                st_trav3 := st_trav3 + DecimalSeparator;
                if Copy(st_trav2, Pos('.', st_trav2) + 1, 255) <> '' then
                    st_trav3 := st_trav3 + Copy(st_trav2, Pos('.', st_trav2) + 1, 255)
                    else
                    st_trav3 := st_trav3 + '0';
                end;
            st_trav2 := st_trav3;
        end;
    end;

begin
    for i_ind1 := 0 to 50 do
        st_types[i_ind1] := ' ';
//  Analyse de la chaine pour la decomposer en items de base
    st_trav1 := Copy(ValeurSeq, 7, Length(ValeurSeq) - 7);
//  Fonction predefinie ???
    if FctPredefinie(st_trav1, NomChampTable, ValeurSeq) then Exit;
//
//    st_trav1 := Copy(st_trav1, 0, Length(st_trav1) - 1);
    st_trav2 := '';
    i_variables := 0;
    for i_chaine := 1 to Length(st_trav1) do
    begin
        st_trav3 := Copy(st_trav1, i_chaine, 1);
        ch_trav1 := st_trav3[1];
        if Pos(ch_trav1, '+-*/') <> 0 then
        begin
            st_types[i_variables] := Determine_Type;
            st_variables[i_variables] := st_trav2;
            st_trav2 := '';
            inc(i_variables);
            st_variables[i_variables] := ch_trav1;
            st_types[i_variables] := 'O';
            inc(i_variables);
        end
        else
        begin
            st_trav2 := st_trav2 + ch_trav1;
        end;
    end;
    st_types[i_variables] := Determine_Type;
    st_variables[i_variables] := st_trav2;
//  Recherche de la nature des champ 'A' : Table destination ou fichier origine ?
    for i_variables := 0 to 50 do
    begin
        if st_types[i_variables] = ' ' then Break;
        if st_types[i_variables] <> 'A' then Continue;
        if Pos(st_variables[i_variables], NomsTable) <> 0 then
        begin
            st_types[i_variables] := 'T';
            Continue;
        end;
        if Pos(st_variables[i_variables], NomsFichier) <> 0 then
            for i_ind1 := 0 to StringGridFichier.RowCount - 1 do
                if StringGridFichier.Cells[0, i_ind1] = st_variables[i_variables] then
                begin
                    st_types[i_variables] := 'F';
                    st_variables[i_variables] := StringGridFichier.Cells[2, i_ind1];
                    Break;
                end;
    end;
//  On peut traiter la formule, le résultat étant mis dans un double
    ValeurSeq := 0.0;
    for i_variables := 0 to 50 do
    begin
        va_temp1 := 0;
        if st_types[i_variables] = ' ' then Break;
        if st_types[i_variables] = 'O' then Continue;
        if st_types[i_variables] = 'T' then
            va_temp1 := TobTable.GetValue(st_variables[i_variables]);
        if st_types[i_variables] = 'F' then
            Extraire_Champ(1, '',st_variables[i_variables], va_temp2, va_temp1);
        if st_types[i_variables] = 'N' then
            va_temp1 := StrToFloat(st_variables[i_variables] + 'E+0');
        if i_variables = 0 then
        begin
            ValeurSeq := va_temp1;
        end
        else
        begin
            if st_variables[i_variables - 1] = '+' then
                ValeurSeq := ValeurSeq + va_temp1;
            if st_variables[i_variables - 1] = '-' then
                ValeurSeq := ValeurSeq - va_temp1;
            if st_variables[i_variables - 1] = '*' then
                ValeurSeq := ValeurSeq * va_temp1;
            if st_variables[i_variables - 1] = '/' then
                if va_temp1 <> 0 then ValeurSeq := ValeurSeq / va_temp1;
        end;
    end;
end;

function  TFRecupNeg.Evaluer_Si(ValeurSeq, NomChampTable : string) : Variant;
var
    TSI, TobSI, TobALORS, TobSINON : TOB;
    Mess, st_trav1 : string;

begin
st_trav1 := V_PGI.DEChamps[TobTable.NumTable, TobTable.GetNumChamp(NomChampTable)].Tipe;
if (st_trav1='INTEGER') or (st_trav1='SMALLINT') or (st_trav1='DOUBLE') or
   (st_trav1='RATE')  then Result:=0                      else
if (st_trav1='DATE')  then Result:=EncodeDate(1900,1,1)   else
if (st_trav1='BLOB') or (st_trav1='DATA') then Result:=#0 else Result:='' ;
TSI := nil;
if not Traite_Si(ValeurSeq, NomsTable, NomsFichier, Mess, TSI) then
    begin
    end;
TobSI    := TSI.FindFirst(['Valeur'], ['SI'], False);
TobALORS := TSI.FindFirst(['Valeur'], ['ALORS'], False);
TobSINON := TSI.FindFirst(['Valeur'], ['SINON'], False);
if Evaluer_Condition(TobSI, NomChampTable) then
    Result := Evaluer_Valeur(TobALORS, NomChampTable)
    else
    if TobSINON <> nil then Result := Evaluer_Valeur(TobSINON, NomChampTable);
TSI.Free;
end;

function  TFRecupNeg.Evaluer_Condition(TSI : TOB; NomChampTable : string) : boolean;
var
    TobTemp : TOB;
    Result_Inter : boolean;
    i_ind1 : integer;
    ValeurSeq, ChampSeq1, ChampSeq2 : Variant;
    st1, Oper_Logique, Oper_Compare, Operande1, Operande2 : string;

begin
Oper_Logique := '';
Oper_Compare := '';
Operande1 := '';
Operande2 := '';
Result := True;
Result_Inter := True;
for i_ind1 := 0 to TSI.Detail.Count - 1 do
    begin
    TobTemp := TSI.Detail[i_ind1];
    if TobTemp.Detail.count <> 0 then
        begin
//  La fille a elle meme des filles il y a donc appelle recursif pour evaluer la condition
//  de niveau inferieur
        if Oper_Logique = '' then
            begin
            Result_Inter := Evaluer_Condition(TobTemp, NomChampTable)
            end
            else
            begin
            if Oper_Logique = 'ET' then
                Result_Inter := (Result_Inter and Evaluer_Condition(TobTemp, NomChampTable))
            else if Oper_Logique = 'OU' then
                Result_Inter := (Result_Inter or Evaluer_Condition(TobTemp, NomChampTable));
            end;
        Continue;
        end;
//  Fille donnant l'operateur logique ET ou OU eventuel
    if TobTemp.GetValue('Type') = 'Logique' then
        begin
        Oper_Logique := Trim(TobTemp.GetValue('Valeur'));
        Continue;
        end
//  Fille donnant l'operateur de comparaison eventuel
    else if TobTemp.GetValue('Type') = 'Operateur' then
        begin
        Oper_Compare := Trim(TobTemp.GetValue('Valeur'));
        Continue;
        end
//  Fille donnant un operande. sur un niveau de condition final il n'y a que 2
//  operande et un operateur.
    else if TobTemp.GetValue('Type') = 'Valeur' then
        begin
        st1 := Trim(TobTemp.GetValue('Valeur'));
        if Operande1 = '' then Operande1 := st1 else Operande2 := st1;
        end;
    end;
//  fin de parcours d'un niveau : si un des operandes est non vide, on recupere la valeur
//  des operandes et on applique l'operateur de comparaison.
//  dans le cas contraire, on retourne Result_Inter.
if (Operande1 = '') and (Operande2 = '') then
    begin
    Result := Result_Inter;
    Exit;
    end;
if not FctPredefinie(Operande1, NomChampTable, ChampSeq1) then
    begin
    if IsNumeric(Operande1) then
        begin
        if Pos(',', Operande1) <> 0 then
            ChampSeq1 := StrToFloat(Operande1)
        else if Pos('.', Operande1) <> 0 then
            ChampSeq1 := StrToFloat(Copy(Operande1, 0, Pos('.', Operande1) - 1) + ',' +
                                    Copy(Operande1, Pos('.', Operande1) + 1, 255))
            else if Operande1 < IntToStr(High(i_ind1)) then
                ChampSeq1 := StrToInt(Operande1)
                else
                ChampSeq1 := Operande1;
        end
    else if Pos(Operande1, NomsTable) <> 0 then
        ChampSeq1 := TobTable.GetValue(Operande1)
    else if Pos(Operande1, NomsFichier) <> 0 then
        Extraire_Champ(0, Operande1, '', ValeurSeq, ChampSeq1)
    else
        begin
        Result := False;
        Exit;
        end;
    end;
if not FctPredefinie(Operande2, NomChampTable, ChampSeq2) then
    begin
    if IsNumeric(Operande2) then
        begin
        if Pos(',', Operande2) <> 0 then
            ChampSeq2 := StrToFloat(Operande2)
        else if Pos('.', Operande2) <> 0 then
            ChampSeq2 := StrToFloat(Copy(Operande2, 0, Pos('.', Operande2) - 1) + ',' +
                                    Copy(Operande2, Pos('.', Operande2) + 1, 255))
            else if Operande2 < IntToStr(High(i_ind1)) then
                ChampSeq2 := StrToInt(Operande2)
                else
                ChampSeq2 := Operande2;
        end
    else if Pos(Operande2, NomsTable) <> 0 then
        ChampSeq2 := TobTable.GetValue(Operande2)
    else if Pos(Operande2, NomsFichier) <> 0 then
        Extraire_Champ(0, Operande2, '', ValeurSeq, ChampSeq2)
    else if Operande2[1] = '"' then ChampSeq2 := Copy(Operande2, 2, Length(Operande2) - 2)
    else
        begin
        Result := False;
        Exit;
        end;
    end;
if VarType(ChampSeq2) = varString then
    begin
    if ChampSeq2 = 'VRAI'      then ChampSeq2 := 'X'
    else if ChampSeq2 = 'FAUX' then ChampSeq2 := '-';
    end;
if Oper_Compare = '='  then Result := (ChampSeq1 = ChampSeq2)
else if Oper_Compare = '<>' then Result := (ChampSeq1 <> ChampSeq2)
else if Oper_Compare = '>'  then Result := (ChampSeq1 > ChampSeq2)
else if Oper_Compare = '>=' then Result := (ChampSeq1 >= ChampSeq2)
else if Oper_Compare = '<'  then Result := (ChampSeq1 < ChampSeq2)
else if Oper_Compare = '<=' then Result := (ChampSeq1 <= ChampSeq2);
end;

function  TFRecupNeg.Evaluer_Valeur(TSI : TOB; NomChampTable : string) : variant;
var
    TobTemp, TobSI, TobALORS, TobSINON : TOB;
    Result_Inter : variant;
    i_ind1 : integer;
    ValeurSeq, ChampSeq1, ChampSeq2 : Variant;
    st1, Operateur, Operande1, Operande2 : string;

begin
Operateur := '';
Operande1 := '';
Operande2 := '';
Result := ' ';
Result_Inter := ' ';
for i_ind1 := 0 to TSI.Detail.Count - 1 do
    begin
    TobTemp := TSI.Detail[i_ind1];
    if TobTemp.Detail.count <> 0 then
        begin
//  La fille a elle meme des filles il y a donc appel recursif pour evaluer la condition
//  de niveau inferieur
        TobSI := TobTemp.FindFirst(['Valeur'], ['SI'], False);
//        if Operateur = '' then
        if TobSI <> nil then
            begin
            TobALORS := TobTemp.FindFirst(['Valeur'], ['ALORS'], False);
            TobSINON := TobTemp.FindFirst(['Valeur'], ['SINON'], False);
            if Evaluer_Condition(TobSI, NomChampTable) then
                Result_Inter := Evaluer_Valeur(TobALORS, NomChampTable)
                else
                if TobSINON <> nil then Result_Inter := Evaluer_Valeur(TobSINON, NomChampTable);
//            Result_Inter := Evaluer_Valeur(TobTemp, NomChampTable)
            end
            else
            begin
//  Fille donnant l'operateur eventuel
            if TobTemp.GetValue('Valeur') = '(' then
                begin
                Result_Inter := Evaluer_Valeur(TobTemp, NomChampTable);
                end
                else
                begin
                if Operateur = '+' then
                    Result_Inter := Result_Inter + Evaluer_Valeur(TobTemp, NomChampTable)
                else if Operateur = '-' then
                    Result_Inter := Result_Inter - Evaluer_Valeur(TobTemp, NomChampTable)
                else if Operateur = '*' then
                    Result_Inter := Result_Inter * Evaluer_Valeur(TobTemp, NomChampTable)
                else if Operateur = '/' then
                    Result_Inter := Result_Inter / Evaluer_Valeur(TobTemp, NomChampTable)
                else if Operateur = '^' then
                    Result_Inter := Power(Result_Inter, Evaluer_Valeur(TobTemp, NomChampTable));
                end;
            end;
        Operande1 := 'Parenthese';
        Continue;
        end;
//  Fille donnant l'operateur eventuel
    if TobTemp.GetValue('Type') = 'Operateur' then
        begin
        Operateur := Trim(TobTemp.GetValue('Valeur'));
        Continue;
        end
//  Fille donnant un operande. sur un niveau de résultat final il n'y a que 2
//  operande et un operateur.
    else if TobTemp.GetValue('Type') = 'Valeur' then
        begin
        st1 := Trim(TobTemp.GetValue('Valeur'));
        if Operande1 = '' then Operande1 := st1 else Operande2 := st1;
//  si l'operateur n'est pas vide, on applique l'operateur aux 2 operandes.
        if Operateur <> '' then
            begin
            if not FctPredefinie(Operande1, NomChampTable, ChampSeq1) then
                begin
                if Operande1 = 'Parenthese' then
                    ChampSeq1 := Result_Inter
                else if IsNumeric(Operande1) then
                    begin
                    if Pos(',', Operande1) <> 0 then
                        ChampSeq1 := StrToFloat(Operande1)
                    else if Pos('.', Operande1) <> 0 then
                        ChampSeq1 := StrToFloat(Copy(Operande1, 0, Pos('.', Operande1) - 1) + ',' +
                                                Copy(Operande1, Pos('.', Operande1) + 1, 255))
                        else if Operande1 < IntToStr(High(i_ind1)) then
                            ChampSeq1 := StrToInt(Operande1)
                            else
                            ChampSeq1 := Operande1;
                    end
                else if Pos(Operande1, NomsTable) <> 0 then
                    ChampSeq1 := TobTable.GetValue(Operande1)
                else if Pos(Operande1, NomsFichier) <> 0 then
                    Extraire_Champ(0, Operande1, '', ValeurSeq, ChampSeq1)
                else if Operande1[1] = '"' then
                    begin
                    ChampSeq1 := Copy(Operande1, 2, Length(Operande1) - 2);
                    if ChampSeq1 = 'VRAI' then ChampSeq1 := 'X';
                    if ChampSeq1 = 'FAUX' then ChampSeq1 := '-';
                    end
                    else
                    begin
                    ChampSeq1 := Operande1;
                    end;
            end;
            if not FctPredefinie(Operande2, NomChampTable, ChampSeq2) then
                begin
                if IsNumeric(Operande2) then
                    begin
                    if Pos(',', Operande2) <> 0 then
                        ChampSeq2 := StrToFloat(Operande2)
                    else if Pos('.', Operande2) <> 0 then
                        ChampSeq2 := StrToFloat(Copy(Operande2, 0, Pos('.', Operande2) - 1) + ',' +
                                                Copy(Operande2, Pos('.', Operande2) + 1, 255))
                        else if Operande2 < IntToStr(High(i_ind1)) then
                            ChampSeq2 := StrToInt(Operande2)
                            else
                            ChampSeq2 := Operande2;
                    end
                else if Pos(Operande2, NomsTable) <> 0 then
                    ChampSeq2 := TobTable.GetValue(Operande2)
                else if Pos(Operande2, NomsFichier) <> 0 then
                    Extraire_Champ(0, Operande2, '', ValeurSeq, ChampSeq2)
                else if Operande2[1] = '"' then
                    ChampSeq2 := Copy(Operande2, 2, Length(Operande2) - 2);
            end;
            if Operateur = '+' then
                begin
                if VarType(ChampSeq1) = varString then
                    begin
                    if Trim(ChampSeq1 + ChampSeq2) = '' then
                        Result_Inter := '"' + ChampSeq1 + ChampSeq2 + '"'
                        else
                        Result_Inter := ChampSeq1 + ChampSeq2;
                    end
                else if VarType(ChampSeq1) = varDouble then
                    Result_Inter := VarAsType(ChampSeq1, varDouble) + VarAsType(ChampSeq2, varDouble);
                end
            else if Operateur = '-' then
                Result_Inter := VarAsType(ChampSeq1, varDouble) - VarAsType(ChampSeq2, varDouble)
            else if Operateur = '*' then
                Result_Inter := VarAsType(ChampSeq1, varDouble) * VarAsType(ChampSeq2, varDouble)
            else if Operateur = '/' then
                Result_Inter := VarAsType(ChampSeq1, varDouble) / VarAsType(ChampSeq2, varDouble)
            else if Operateur = '^' then
                Result_Inter := Power(VarAsType(ChampSeq1, varDouble), VarAsType(ChampSeq2, varDouble));
            Operande1 := Result_Inter;
            end;
        end;
    end;
if not FctPredefinie(Operande1, NomChampTable, ChampSeq1) then
    begin
    if Operande1 = 'Parenthese' then
        ChampSeq1 := Result_Inter
    else if IsNumeric(Operande1) then
        begin
        if Pos(',', Operande1) <> 0 then
            ChampSeq1 := StrToFloat(Operande1)
        else if Pos('.', Operande1) <> 0 then
            ChampSeq1 := StrToFloat(Copy(Operande1, 0, Pos('.', Operande1) - 1) + ',' +
                                    Copy(Operande1, Pos('.', Operande1) + 1, 255))
            else if Operande1 < IntToStr(High(i_ind1)) then
                ChampSeq1 := StrToInt(Operande1)
                else
                ChampSeq1 := Operande1;
        end
    else if Pos(Operande1, NomsTable) <> 0 then
        ChampSeq1 := TobTable.GetValue(Operande1)
    else if Pos(Operande1, NomsFichier) <> 0 then
        Extraire_Champ(0, Operande1, '', ValeurSeq, ChampSeq1)
    else if Operande1[1] = '"' then
        begin
        ChampSeq1 := Copy(Operande1, 2, Length(Operande1) - 2);
        if ChampSeq1 = 'VRAI' then ChampSeq1 := 'X';
        if ChampSeq1 = 'FAUX' then ChampSeq1 := '-';
        end;
    end;
if Operateur = '' then
    Result := ChampSeq1
    else
    Result := Result_Inter;
end;

function  TFRecupNeg.FctPredefinie(st_trav1, NomChampTable : string; var ValeurSeq : Variant) : boolean;
var
    st_trav2 : string;

begin
Result := False;
//  Verification de la presence d'une fonction prédéfinie
if Pos('MAJUSCULE(', st_trav1) <> 0 then
    begin
    st_trav2 := Copy(st_trav1, Pos('(', st_trav1) + 1, Length(st_trav1) - Pos('(', st_trav1) - 1);
    ValeurSeq := Fct_MAJUSCULE(st_trav2);
    Result := True;
    Exit;
    end;
if (Pos('COMPTEUR', st_trav1) <> 0) and (Length(st_trav1) = 8) then
    begin
    ValeurSeq := Fct_COMPTEUR(NomChampTable);
    Result := True;
    Exit;
    end;
if Pos('REPLACECAR(', st_trav1) <> 0 then
    begin
    st_trav2 := Copy(st_trav1, Pos('(', st_trav1) + 1, Length(st_trav1) - Pos('(', st_trav1) - 1);
    ValeurSeq := Fct_REPLACECAR(st_trav2);
    st_trav1 := Copy(NomChampTable, Pos('_', NomChampTable) + 1, 255);
    if st_trav2 <> ValeurSeq then
        begin
        st_trav1 := st_trav1 + ';' + st_trav2 + ';' + ValeurSeq;
        if Correspond.IndexOf(st_trav1) < 0 then
            begin
            writeln(Corresp, st_trav1);
            flush (Corresp);
            Correspond.Add(st_trav1);
            end;
        end;
    Result := True;
    Exit;
    end;
if Pos('CORRESPOND(', st_trav1) <> 0 then
    begin
    st_trav2 := Copy(ValeurSeq, Pos('(', ValeurSeq) + 1, Length(ValeurSeq) - Pos('(', ValeurSeq) - 2);
    if st_trav2 = '' then
        st_trav2 := Copy(st_trav1, Pos('(', st_trav1) + 1, Length(st_trav1) - Pos('(', st_trav1) - 1);
    ValeurSeq := Fct_CORRESPOND(NomChampTable, st_trav2);
    Result := True;
    Exit;
    end;
if Pos('TABLE(', st_trav1) <> 0 then
    begin
    st_trav2 := Copy(ValeurSeq, Pos('(', ValeurSeq) + 1, Length(ValeurSeq) - Pos('(', ValeurSeq) - 2);
    if st_trav2 = '' then
        st_trav2 := Copy(st_trav1, Pos('(', st_trav1) + 1, Length(st_trav1) - Pos('(', st_trav1) - 1);
    ValeurSeq := Fct_TABLE(NomChampTable, st_trav2);
    Result := True;
    Exit;
    end;
if Pos('SOUSCHAINE(', st_trav1) <> 0 then
    begin
    st_trav2 := Copy(ValeurSeq, Pos('(', ValeurSeq) + 1, Length(ValeurSeq) - Pos('(', ValeurSeq) - 2);
    if st_trav2 = '' then
        st_trav2 := Copy(st_trav1, Pos('(', st_trav1) + 1, Length(st_trav1) - Pos('(', st_trav1) - 1);
    ValeurSeq := Fct_SOUSCHAINE(st_trav2);
    Result := True;
    Exit;
    end;
if Pos('SPACE(', st_trav1) <> 0 then
    begin
    st_trav2 := Copy(ValeurSeq, Pos('(', ValeurSeq) + 1, Length(ValeurSeq) - Pos('(', ValeurSeq) - 2);
    if st_trav2 = '' then
        st_trav2 := Copy(st_trav1, Pos('(', st_trav1) + 1, Length(st_trav1) - Pos('(', st_trav1) - 1);
    ValeurSeq := Fct_SPACE(st_trav2);
    Result := True;
    Exit;
    end;
if Pos('STRINT(', st_trav1) <> 0 then
    begin
    st_trav2 := Copy(ValeurSeq, Pos('(', ValeurSeq) + 1, Length(ValeurSeq) - Pos('(', ValeurSeq) - 2);
    if st_trav2 = '' then
        st_trav2 := Copy(st_trav1, Pos('(', st_trav1) + 1, Length(st_trav1) - Pos('(', st_trav1) - 1);
    ValeurSeq := Fct_STRINT(st_trav2);
    Result := True;
    Exit;
    end;
if Pos('INTSTR(', st_trav1) <> 0 then
    begin
    st_trav2 := Copy(ValeurSeq, Pos('(', ValeurSeq) + 1, Length(ValeurSeq) - Pos('(', ValeurSeq) - 2);
    if st_trav2 = '' then
        st_trav2 := Copy(st_trav1, Pos('(', st_trav1) + 1, Length(st_trav1) - Pos('(', st_trav1) - 1);
    ValeurSeq := Fct_INTSTR(st_trav2);
    Result := True;
    Exit;
    end;
if Pos('DATE(', st_trav1) <> 0 then
    begin
    st_trav2 := Copy(ValeurSeq, Pos('(', ValeurSeq) + 1, Length(ValeurSeq) - Pos('(', ValeurSeq) - 2);
    if st_trav2 = '' then
        st_trav2 := Copy(st_trav1, Pos('(', st_trav1) + 1, Length(st_trav1) - Pos('(', st_trav1) - 1);
    ValeurSeq := Fct_DATE(st_trav2);
    Result := True;
    Exit;
    end;
end;

function  TFRecupNeg.Fct_MAJUSCULE(var Syntaxe : string) : variant;
var
    ValeurSeq : variant;

begin
ValeurSeq := ' ';
ValeurSeq := RecupValeur(Syntaxe);
if VarType(ValeurSeq) = varString then
    Result := UpperCase(ValeurSeq)
    else
    Result := ValeurSeq;
end;

function  TFRecupNeg.Fct_COMPTEUR(NomChampTable : string) : variant;
var
    Q : TQuery;

begin
if TOBCompteur.FieldExists(NomChampTable) then
   Result := TOBCompteur.GetValue(NomChampTable) + 1
   else
   begin
//   Result := 1;
   Q := OpenSql('Select Max(' + NomChampTable + ') From ' + TobTable.NomTable,True);
   FetchSQLODBC(Q);
   If Not Q.Eof Then Result := Q.Fields[0].AsInteger;
   Ferme(Q);
   Inc(Result);
   TOBCompteur.AddChampSup(NomChampTable, False);
   end;
TOBCompteur.PutValue(NomChampTable,Result);
end;

function  TFRecupNeg.Fct_REPLACECAR(var Syntaxe : string) : Variant;
var
    Valeur, Remplace, Remplacant, st1 : string;
    ValeurSeq, ChampSeq : variant;

begin
Valeur := ReadTokenPipe(Syntaxe, ',');
Remplace := ReadTokenPipe(Syntaxe, ',');
Remplacant := ReadTokenPipe(Syntaxe, ',');
ValeurSeq := ' ';
ChampSeq := RecupValeur(Valeur);
Syntaxe := ChampSeq;
st1 := ChampSeq;
while Pos(Remplace, st1) > 0 do
    st1 := Copy(st1, 0, Pos(Remplace, st1) - 1) + Remplacant +
           Copy(st1, Pos(Remplace, st1) + Length(Remplace), 255);
Result := st1;
end;

function  TFRecupNeg.Fct_CORRESPOND(NomChampTable, NomChampSeq : string) : Variant;
var
    i_ind1 : integer;
    ValeurSeq, ChampSeq : variant;
    st_trav1, st_trav2, st_trav3, Valeur, Nom : string;

begin
ValeurSeq := ' ';
Extraire_Champ(0, NomChampSeq, '', ValeurSeq, ChampSeq);
Result := ChampSeq;
Valeur := ChampSeq;
Nom := Copy(NomChampTable, Pos('_', NomChampTable) + 1, 255);
for i_ind1 := 0 to Correspond.Count - 1 do
    begin
    st_trav3 := Correspond.Strings[i_ind1];
    st_trav1 := ReadTokenSt(st_trav3);
    st_trav2 := ReadTokenSt(st_trav3);
    if (Nom = st_trav1) and (Valeur = st_trav2) then
        begin
        Result := st_trav3;
        Exit;
        end;
    end;
end;

function  TFRecupNeg.Fct_TABLE(NomChampTable, NomChampSeq : string) : Variant;
var
    TableALire, TableEnCours : TOB;
    ValeurSeq : variant;
    st_trav1, st_trav2, st_trav3, st_sep, Valeur, Select : string;
    Q : TQuery;

begin
ValeurSeq := ' ';
TableALire   := Tob_DeTables.FindFirst(['DT_NOMTABLE'], [NomChampSeq], False);
TableEnCours := Tob_DeTables.FindFirst(['DT_NOMTABLE'], [TobTable.NomTable], False);
if TOBArray[TableToNum(NomChampSeq)] <> nil then
    begin
    TOBALire := TOB.Create('', nil, -1);
    TOBALire.Dupliquer(TOBArray[TableToNum(NomChampSeq)],True,True);
    end
    else
    TOBAlire := nil;
if TOBALire <> nil then
    begin
    st_trav1 := TableALire.GetValue('DT_CLE1');
    Valeur := '';
    st_sep := '';
    while st_trav1 <> '' do
        begin
        st_trav2 := ReadTokenStV(st_trav1);
        st_trav3 := Copy(st_trav2, Pos('_', st_trav2), 255);
        st_trav3 := TableEnCours.GetValue('DT_PREFIXE') + st_trav3;
        Valeur := Valeur + st_sep + st_trav2 + '=';
        ValeurSeq := TobTable.GetValue(st_trav3);
        Case VarType(ValeurSeq) of
        varSmallint : Valeur := Valeur + IntToStr(ValeurSeq);
        varInteger  : Valeur := Valeur + IntToStr(ValeurSeq);
        varSingle   : Valeur := Valeur + IntToStr(ValeurSeq);
        varDouble   : Valeur := Valeur + FloatToStr(ValeurSeq);
        varDate     : Valeur := Valeur + UsDateTime(ValeurSeq);
        varString   : Valeur := Valeur + '"' + ValeurSeq + '"';
        end;
        st_sep := ' and ';
        end;
    if (TOBALire.NomTable = NomChampSeq) and (Valeur = CleArray[TOBALire.NumTable]) then
        begin
        Valeur := TableALire.GetValue('DT_PREFIXE') + Copy(NomChampTable, Pos('_', NomChampTable), 255);
        Result := TOBALire.GetValue(Valeur);
        TOBALire.Free; TOBALire := nil;
        Exit;
        end;
    TOBALire.Free; TOBALire := nil;
    end;
TOBALire := TOB.Create(NomChampSeq, nil, -1);
TOBALire.InitValeurs;
st_trav1 := TableALire.GetValue('DT_CLE1');
Select := 'Select * from ' + NomChampSeq + ' where ';
Valeur := '';
st_sep := '';
while st_trav1 <> '' do
    begin
    st_trav2 := ReadTokenStV(st_trav1);
    st_trav3 := Copy(st_trav2, Pos('_', st_trav2), 255);
    st_trav3 := TableEnCours.GetValue('DT_PREFIXE') + st_trav3;
    Valeur := Valeur + st_sep + st_trav2 + '=';
    ValeurSeq := TobTable.GetValue(st_trav3);
    Case VarType(ValeurSeq) of
    varSmallint : Valeur := Valeur + IntToStr(ValeurSeq);
    varInteger  : Valeur := Valeur + IntToStr(ValeurSeq);
    varSingle   : Valeur := Valeur + IntToStr(ValeurSeq);
    varDouble   : Valeur := Valeur + FloatToStr(ValeurSeq);
    varDate     : Valeur := Valeur + UsDateTime(ValeurSeq);
    varString   : Valeur := Valeur + '"' + ValeurSeq + '"';
    end;
    st_sep := ' and ';
    end;
st_trav1 := V_PGI.DEChamps[TobTable.NumTable, TobTable.GetNumChamp(NomChampTable)].Tipe;
if (st_trav1='INTEGER') or (st_trav1='SMALLINT') or (st_trav1='DOUBLE') or
   (st_trav1='RATE')  then Result:=0                      else
if (st_trav1='DATE')  then Result:=EncodeDate(1900,1,1)   else
if (st_trav1='BLOB') or (st_trav1='DATA') then Result:=#0 else Result:='' ;
Select := Select + Valeur;
Q := OpenSQL(Select, True,-1,'',true);
if not Q.Eof then
    begin
    TOBALire.SelectDB('', Q);
    st_trav1 := TableALire.GetValue('DT_PREFIXE') + Copy(NomChampTable, Pos('_', NomChampTable), 255);
    Result := TOBALire.GetValue(st_trav1);
    CleArray[TableToNum(NomChampSeq)] := Valeur;
    TOBArray[TableToNum(NomChampSeq)] := TOB.Create('', nil, -1);
    TOBArray[TableToNum(NomChampSeq)].Dupliquer(TOBALire,True,True);
    end;
Ferme(Q);
TOBALire.Free; TOBALire := nil;
end;

function  TFRecupNeg.Fct_SOUSCHAINE(Syntaxe : string) : Variant;
var
    Valeur, st1 : string;
    Debut, Longueur : integer;
    ChampSeq : variant;

begin
Valeur := ReadTokenPipe(Syntaxe, ',');
Debut := StrToInt(ReadTokenPipe(Syntaxe, ','));
Longueur := StrToInt(ReadTokenPipe(Syntaxe, ','));
ChampSeq := RecupValeur(Valeur);
st1 := ChampSeq;
Result := Copy(st1, Debut, Longueur);
end;

function  TFRecupNeg.Fct_SPACE(Syntaxe : string) : Variant;
var
    i_ind1, Longueur : integer;

begin
Longueur := StrToInt(Syntaxe);
for i_ind1 := 1 to Longueur do Result := Result + ' ';
end;

function  TFRecupNeg.Fct_INTSTR(Syntaxe : string) : Variant;
begin
Result := RecupValeur(Syntaxe);
end;

function  TFRecupNeg.Fct_STRINT(Syntaxe : string) : Variant;
begin
Result := RecupValeur(Syntaxe);
end;

function  TFRecupNeg.Fct_DATE(Syntaxe : string) : Variant;
begin
Result := StrToDate(Syntaxe);
end;

function  TFRecupNeg.RecupValeur(Syntaxe : string) : Variant;
var
    Valeur, st1 : string;
    ValeurSeq, ChampSeq : variant;

begin
Result := '';
if Pos('+', Syntaxe) = 0 then
    begin
    Extraire_Champ(0, Syntaxe, '', ValeurSeq, ChampSeq);
    Result := ChampSeq;
    end
    else
    begin
    Valeur := Syntaxe;
    while Valeur <> '' do
        begin
        st1 := ReadTokenPipe(Valeur, '+');
        Extraire_Champ(0, st1, '', ValeurSeq, ChampSeq);
        if VarType(ChampSeq) = varString then
            Result := Result + ChampSeq
            else
            Result := ChampSeq;
        end;
    end;
end;

procedure TFRecupNeg.Prepare_Dispo(var traite : boolean);
var
    ficdesori, ficdesdes, ficdataori, ficdatades : textfile;
    LastOffset, i_offset, i_longueur, i_ind1, i_ind2, NbMois : integer;
    st1, st2, stdesori, stdesdes, stdessav : string;
    StkPhy : double;
    FinMois : TDateTime;
    DateDebut : string;
    ES : array [1..24] of double;
    ValeurSeq, ChampSeq : variant;

begin
//  le fichier DISPO est a traiter. donc on va le pretraiter.
//  format attendu dans le .des : des champs dans l'ordre que l'on veut, sauf en fin de liste
//  les 12 mois entrees et sorties qui se suivent dans l'ordre E1, E2,...,E12,S1,S2,...,S12
//  le traitement ci dessous se base sur le nom de champ STO0450 pour les 12 entrees et
//  les 12 sorties, et sur le champ STO0400 pour le stock initial, situé avant les
//  entrees et les sorties.
Traite := True;
for i_ind1 := 0 to StringGridFichier.RowCount - 1 do
    st1 := st1 + StringGridFichier.Cells[0, i_ind1] + ';';
if (Pos('STO0400', st1) = 0) or (Pos('STO0450', st1) = 0) then
    begin
    Traite := False;
    Exit;
    end;
//  preparation du dictionnaire modifié
st1 := StringGrid1.Cells[0, 0] + '.des';
st2 := StringGrid1.Cells[0, 0] + '_des.sai';
RenameFile(st1, st2);
AssignFile (ficdesori, st2);
Reset(ficdesori);
readln(ficdesori, stdesori);
AssignFile(ficdesdes, st1);
Rewrite(ficdesdes);
LastOffset := 0;
while stdesori <> '' do
    begin
    stdesdes := stdesori;
    if Pos('STO0400', stdesori) <> 0 then
        begin
        stdesdes := Copy(stdesori, 0, 39) + IntToStr(LastOffset) + Copy(stdesori, 44, 255);
        Writeln(ficdesdes, stdesdes);
        Writeln(ficdesdes, 'STO0999;Date fin de mois              ;' +
                           IntToStr(LastOffset+15) + ';010;0;D;       ');
        st_DateDispo := IntToStr(LastOffset+15) + ';010;0;D;';
        Break;
        end
        else
        begin
        Writeln(ficdesdes, stdesdes);
        st2 := stdesdes;
        st1 := ReadTokenSt(st2);
        st1 := ReadTokenSt(st2);
        st1 := ReadTokenSt(st2);
        LastOffset := LastOffset + StrToInt(ReadTokenSt(st2));
        end;
    readln(ficdesori, stdesori);
    end;
CloseFile(ficdesori);
CloseFile(ficdesdes);
//  Chargement du dictionnaire d'origine
st1 := StringGrid1.Cells[0, 0] + '_des.sai';
AssignFile (ficdesori, st1);
Reset(ficdesori);
for i_ind1 := 0 to StringGridFichier.RowCount - 1 do StringGridFichier.Rows[i_ind1].Clear;
i_ind1 := 0;
readln(ficdesori, stdesori);
while stdesori <> '' do
begin
    st1 := Trim(Copy(stdesori, 1, Pos(';',stdesori) - 1));
    st2 := Trim(Copy(stdesori, Pos(';',stdesori)+1, 255));
    StringGridFichier.Cells[0, i_ind1] := st1;
    stdesori := st2;
    st1 := Trim(Copy(stdesori, 1, Pos(';',stdesori) - 1));
    st2 := Trim(Copy(stdesori, Pos(';',stdesori)+1, 255));
    StringGridFichier.Cells[1, i_ind1] := st1;
    stdesori := st2;
    StringGridFichier.Cells[2, i_ind1] := stdesori;
    readln(ficdesori, stdesori);
    i_ind1 := Succ(i_ind1);
end;
//  traitement des données, lues sur le format d'origine et ecrites sur le format modifie
st1 := StringGrid1.Cells[0, 0] + '.pgi';
st2 := StringGrid1.Cells[0, 0] + '_pgi.sai';
RenameFile(st1, st2);
AssignFile (ficdataori, st2);
Reset(ficdataori);
readln(ficdataori, st_enreg);
AssignFile(ficdatades, st1);
Rewrite(ficdatades);
while st_enreg <> '' do
    begin
    stdesdes := '';
    for i_ind1 := 0 to StringGridFichier.RowCount - 1 do
        begin
        if StringGridFichier.Cells[0, i_ind1] = '' then Break;
        if Pos('STO0400', StringGridFichier.Cells[0, i_ind1]) = 0 then
            begin
            st1 := StringGridFichier.Cells[2, i_ind1];
            i_offset := StrToInt(ReadTokenSt(st1));
            i_longueur := StrToInt(ReadTokenSt(st1));
            stdesdes := stdesdes + Copy(st_enreg, i_offset, i_longueur);
            end
            else
            begin
            DateDebut := ParamGenData.Strings[ParamGenCle.IndexOf('DebutExercice')];
            NbMois := StrToInt(ParamGenData.Strings[ParamGenCle.IndexOf('NbMois')]);
//  on sauvegarde le debut d'enregistrement qui va etre inchangé dans la suite
            stdessav := stdesdes;
//  on extrait le stock initial du champ STO0400
            st1 := StringGridFichier.Cells[2, i_ind1];
            Extraire_Champ(1, '', st1, ValeurSeq, ChampSeq);
            StkPhy := ChampSeq;
//  on extrait ensuite les 12 mois entree et les 12 mois sortie
            for i_ind2 := i_ind1 + 1 to 24 do
                begin
                st1 := StringGridFichier.Cells[2, i_ind2];
                Extraire_Champ(1, '', st1, ValeurSeq, ChampSeq);
                ES[i_ind2 - i_ind1] := ChampSeq;
                end;
//  on ecrit un enregistrement par mois en calculant le stock en fin de mois
            for i_ind2 := 1 to NbMois do
                begin
                StkPhy := StkPhy + ES[i_ind2] - ES[i_ind2 + 12];
                FinMois := FinDeMois(PlusMois(StrToDate(DateDebut), i_ind2 - 1));
                st1 := Format('%14.3f', [StkPhy]);
                if Pos('-', st1) <> 0 then
                    begin
                    st1[Pos('-', st1)] := ' ';
                    st1[1] := '-';
                    end;
                while Pos(' ', st1) <> 0 do st1[Pos(' ', st1)] := '0';
                while Pos(',', st1) <> 0 do st1[Pos(',', st1)] := '.';
                if st1[1] = '0' then st1[1] := ' ';
                stdesdes := stdessav + st1 + DateToStr(FinMois);
                Writeln(ficdatades, stdesdes);
                end;
            st1 := Format('%14.3f', [StkPhy]);
            if Pos('-', st1) <> 0 then
                begin
                st1[Pos('-', st1)] := ' ';
                st1[1] := '-';
                end;
            while Pos(' ', st1) <> 0 do st1[Pos(' ', st1)] := '0';
            while Pos(',', st1) <> 0 do st1[Pos(',', st1)] := '.';
            if st1[1] = '0' then st1[1] := ' ';
            stdesdes := stdessav + st1 + '01/01/1900';
            Writeln(ficdatades, stdesdes);
            end;
        end;
    readln(ficdataori, st_enreg);
    end;
CloseFile(ficdataori);
CloseFile(ficdatades);

end;

procedure TFRecupNeg.GereConvertError(NomChampTable : string; ValeurSeq : Variant);
var
   st_log, st_trav1, st_trav2 : string;

begin
if NomChampTable = '' then Exit;
st_log := 'Erreur : Ligne ' + IntToStr(Nb_Lus) + ' lue, champ ' + NomChampTable ;
st_log := st_log + ', valeur : ';
st_trav1 := V_PGI.DEChamps[TobTable.NumTable, TobTable.GetNumChamp(NomChampTable)].Tipe;
st_trav2 := VarAsType(ValeurSeq, varString);
st_log := st_log + st_trav2 + ', attendu : ' + st_trav1;
writeln (ErreurLog, st_log);
flush (ErreurLog);
end;

procedure TFRecupNeg.GereEngineError(NomTable, SQL : string);
var
   st_trav1 : string;

begin
if NomTable = '' then Exit;
st_trav1 := 'Erreur d''accés à la table ' + NomTable;
writeln (ErreurLog, st_trav1);
flush (ErreurLog);
writeln (ErreurLog, SQL);
flush (ErreurLog);
end;

procedure TFRecupNeg.VerifieSQL(var Sql : string);
begin
if Pos('''', Sql) <> 0 then
    while Pos(':', Sql) <> 0 do
        Sql[Pos(':', Sql)] := '=';
end;

Initialization

end.


