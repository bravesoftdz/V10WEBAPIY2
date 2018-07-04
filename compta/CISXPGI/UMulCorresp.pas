unit UMulCorresp;

interface

uses
  Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, HSysMenu, Menus, Db, Hqry, HRichOLE,
  HRichEdt, Grids, HDB, StdCtrls, HTB97, Hctrls, HEnt1, HPanel,
  Variants, UIUtil, hmsgbox, HStatus, LicUtil, ColMemo,
  DBGrids, ComCtrls, ExtCtrls, uDbxDataSet, ADODB,
  Mask;

type
  TFCorresp = class(TFMul)
    QProfile: TStringField;
    QCodedepart: TStringField;
    QCodearrive: TStringField;
    QNature: TStringField;
    BDelete: TToolbarButton97;
    QChampcode: TStringField;
    QChampnature: TStringField;
    QCommentaire: TStringField;
    Label4: TLabel;
    Comment: TEdit;
    Label2: TLabel;
    Profil: TEdit;
    QDomaine: TStringField;
    Breprise: TToolbarButton97;
    POPZ: TPopupMenu;
    Partable1: TMenuItem;
    Enliste1: TMenuItem;
    QIdent: TIntegerField;
    BCherche1: TToolbarButton97;
    QTablename: TStringField;
    procedure BChercheClick(Sender: TObject); override;
    procedure FormShow(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject); override;
    procedure BrepriseClick(Sender: TObject);
    procedure Partable1Click(Sender: TObject);
    procedure Enliste1Click(Sender: TObject);
    procedure BinsertClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
  private
    { Déclarations privées }
    procedure SuppressionCorresp(Ident : integer; Name, profile : string);
    procedure SaveListeCorresp(Ask : Boolean);
  public
    { Déclarations publiques }
  end;

var
  FCorresp: TFCorresp;



procedure MulticritereCorresp (PP: THPanel) ;

implementation

uses UDMIMP
,UCreatCorresp
,UPDomaine
,UCorresp
,Uscript
;

{$R *.DFM}

procedure MulticritereCorresp (PP: THPanel) ;
begin
 FCorresp := TFCorresp.Create(Application);
 InitInside(FCorresp, PP);
 FCorresp.Show;
end;


procedure TFCorresp.BChercheClick(Sender: TObject);
var
  SelectQL : string;
  SavePlace: TBookmark;
  St       : string;
begin
  SelectQL := ' SELECT * from '+ DMImport.GzImpCorresp.TableName;
  Q.Close;
  Q.SQL.Clear;

  St := ' Where Tablename <> "" ';
  if Comment.Text <> ''  then St := St + ' and Commentaire like "%' + Comment.Text + '%" ';
  if profil.Text <> ''   then St := St + ' and Profile like "%' + profil.Text + '%" ';
//  if Code.Text <> ''     then St := St + ' and Tablename like "%' + Code.Text + '%" ';
  if  (ParamCount > 0) and (GetInfoVHCX.Mode = 'I') and (GetInfoVHCX.Domaine <> '') then
  begin
      St := St + ' and (Domaine = "" or Domaine="'+ GetInfoVHCX.Domaine+'")';
  end;

  Q.SQL.Add(SelectQL + St + ' Order By Domaine,Ident,TableName,Profile');
  Q.SqlConnection := DMImport.Db ; // XP 07.01.2008
  ChangeSQL(Q);
  Q.Open;
  Q.UpdateCumul(PCumul, FListe);
  Z_SQL.Lines.Clear;
  Z_SQL.Lines.Add(Q.SQL[0]);
  if (Sender=Nil) and Not((Q.EOF) and (Q.BOF)) then SavePlace:=Q.GetBookmark else SavePlace:=Nil ;
  if (Sender=Nil) and Not((Q.EOF) and (Q.BOF)) then if SavePlace<>Nil then
   BEGIN
       Try
       Q.GotoBookmark(SavePlace) ;
       Q.FreeBookmark(SavePlace) ;
       Except
       End ;
   END ;
  HMTrad.ResizeDBGridColumns(FListe);
  CentreDBGrid(FListe);
end;


procedure TFCorresp.FormShow(Sender: TObject);
begin
  inherited;
            if V_PGI.PassWord = DayPass(Date) then
               Binsert.DropDownMenu := POPZ
            else
               Binsert.Onclick := BinsertClick;
end;

procedure TFCorresp.SuppressionCorresp (Ident : integer; Name, profile : string);
var
  tbcharger: TADOTable;
begin
  tbCharger := TADOTable.Create(Application);
  tbCharger.Connection := DMImport.Db as TADOConnection;
  tbCharger.TableName := DMImport.GzImpCorresp.TableName;
  tbCharger.open;

  try
    if tbCharger.Locate('Ident;TableName;Profile',VarArrayOf([Ident,Name,profile]), [loCaseInsensitive]) then
    begin
      tbCharger.Delete;
    end;
  except
    PGIBox(Exception(ExceptObject).message, '');
    exit;
  end;
  tbcharger.Close;  tbcharger.free;
end;

procedure TFCorresp.BDeleteClick(Sender: TObject);
var
i     : integer;
Texte : string;
begin
  inherited;
  if (Fliste.NbSelected=0) and (not Fliste.AllSelected) then
  begin
    MessageAlerte('Aucun élément sélectionné');
    exit;
  end;
  Texte:='Vous allez supprimer définitivement les Scripts.#13#10Confirmez vous l''opération ?';
  if HShowMessage('0;Suppression;'+Texte+';Q;YN;N;N;','','')<>mrYes then exit ;
  if Fliste.AllSelected then
  BEGIN
    Q.First;
    while Not Q.EOF do
    BEGIN
      MoveCur(False);
      SuppressionCorresp(Q.FindField('Ident').Asinteger,Q.FindField('Tablename').AsString,Q.FindField('Profile').AsString) ;
      Q.Next;
    END;
    Fliste.AllSelected:=False;
  END
  ELSE
  BEGIN
    InitMove(Fliste.NbSelected,'');
    for i:=0 to Fliste.NbSelected-1 do
    BEGIN
      MoveCur(False);
      Fliste.GotoLeBookmark(i);
      SuppressionCorresp(Q.FindField('Ident').Asinteger,Q.FindField('Tablename').AsString,Q.FindField('Profile').AsString) ;
    END;
    Fliste.ClearSelected;
  END;
  FiniMove;
  BChercheClick(Sender);
end;

procedure TFCorresp.FListeDblClick(Sender: TObject);
var
i                     : integer;
FCor                  : PListeCorresp;
TCorr                 : TADOTable;
AStreamTable          : TmemoryStream;
ATblCorField          : TField;
ScriptCorr            : TScriptCorresp;
begin
  inherited;
    ScriptCorr := nil;
    if Q.FindField('Tablename').AsString = TraduireMemoire('Liste de correspondance') then
    begin
                 TCorr := TADOTable.Create(Application);
                 TCorr.Connection := DMImport.Db as TADOConnection;
                 TCorr.TableName := DMImport.GzImpCorresp.TableName;

                 CorrespDlg := TCorrespDlg.Create(nil);
                 CorrespDlg.Script := nil;
                 CorrespDlg.Champ := nil;
                 try
                   with TCorr do
                   begin
                        Open;
                        if Locate('Ident;Tablename;Profile;Champcode',VarArrayOf([Q.FindField('Ident').asinteger, Q.FindField('Tablename').AsString, Q.FindField('Profile').Asstring, Q.FindField('Champcode').Asstring]), [loCaseInsensitive]) then
                        begin

                              CorrespDlg.Profile.Text :=  FieldByName('Profile').Asstring;
                              CorrespDlg.Famille.Text := FieldByName('Champcode').Asstring;
                              CorrespDlg.Memo1.text := FieldByName('Commentaire').Asstring;
                              CorrespDlg.Domaine.itemindex  := CorrespDlg.Domaine.items.IndexOf(RendLibelleDomaine(Q.FindField('Domaine').Asstring));
                              ATblCorField := FieldByName('TBLCOR');
                              AStreamTable := TmemoryStream.create;
                              TBlobField(ATblCorField).SaveToStream (AStreamTable);
                              AStreamTable.Seek (0,0);
                              ScriptCorr := LoadScriptCorresp(AStreamTable);
                  //TMemoField(ATblCorField).SaveToFile (CurrentDossier+'xxxxx.txt');
                              With ScriptCorr do
                              begin
                                  CorrespDlg.Profile.Text := ScriptCorr.Name;
                                  CorrespDlg.edNomTable.Text := ScriptCorr.FFichier;
                                  CorrespDlg.StringGrid1.RowCount := LFEntree.Count+1;
                                  for i:=0 to LFEntree.Count-1 do
                                  begin
                                       CorrespDlg.Table.FEntree.Add(LFEntree[i]);
                                       CorrespDlg.Table.FSortie.add(LFSortie[i]);
		                       CorrespDlg.StringGrid1.Cells[0, i] := LFEntree[i];
		                       CorrespDlg.StringGrid1.Cells[1, i] := LFSortie[i];
                                  end;
                              end;
                              AStreamTable.free;
                              Close;
                              if CorrespDlg.ShowModal = mrOK then
                               SaveListeCorresp( FALSE);
                        end;
                   end;
                 finally
                         ScriptCorr.destroy;
                         CorrespDlg.Libere;
                         CorrespDlg.Free;
                 end;
                 exit;
    end;
    FCreatCorresp := TFCreatCorresp.Create(Application);

    With FCreatCorresp do
    begin

          Try
                 Ident.value       := Q.FindField('Ident').Asinteger;
                 Profile.Text      := Q.FindField('Profile').AsString;
                 ChampCode.Text    := Q.FindField('Champcode').AsString;
                 Codedepart.Text   := Q.FindField('Codedepart').AsString;
                 Coderemplace.Text := Q.FindField('Codearrive').AsString;
                 Champnature.Text  := Q.FindField('Champnature').AsString;
                 Commentaire.Text  := Q.FindField('Commentaire').Asstring;
                 Domaine.itemindex  := Domaine.items.IndexOf(RendLibelleDomaine(Q.FindField('Domaine').Asstring));
                 ChargementComboparam (TableN, Q.FindField('Domaine').Asstring);
                 TableN.value := Q.FindField('Tablename').AsString;
                 ChargeNature(Q.FindField('Nature').AsString) ;
                 Nature.itemindex  := Nature.items.Indexof(Q.FindField('Nature').AsString);
                 ShowModal ;
          Finally
                 TCorr := TADOTable.Create(Application);
                 TCorr.Connection := DMImport.Db as TADOConnection;
                 TCorr.TableName := DMImport.GzImpCorresp.TableName;

                 for i:=0 To  FCreatCorresp.ListeInfoCorresp.Count-1 do
                 begin
                   FCor :=ListeInfoCorresp.Items[i];
                   with TCorr do
                   begin
                        if Active then Close;
                        Open;
                        if Locate('Ident;Tablename;Profile',VarArrayOf([FCor^.Ident, FCor^.Tablename, FCor^.Profile]), [loCaseInsensitive]) then Edit
                        else Insert;
                        FieldByName('Ident').AsInteger          := FCor^.Ident;
                        FieldByName('Tablename').Asstring       := FCor^.Tablename;
                        FieldByName('Profile').Asstring         := FCor^.Profile;
                        FieldByName('Champcode').Asstring       := FCor^.Champcode;
                        FieldByName('Codedepart').Asstring      := FCor^.Codedepart;
                        FieldByName('Codearrive').Asstring      := FCor^.Codedearrive;
                        FieldByName('Champnature').Asstring     := FCor^.Champnature;
                        FieldByName('Nature').Asstring          := FCor^.Nature;
                        FieldByName('Commentaire').Asstring     := FCor^.Commentaire;
                        FieldByName('Domaine').Asstring         := RendDomaine(FCor^.Domaine);
                        Post;
                        Close;
                   end;

                 end;
                 freeListe;
                 Free ;
                 TCorr.free;
           End;
     end;
     BChercheClick(Sender);
end;

procedure TFCorresp.BrepriseClick(Sender: TObject);
{$IFNDEF DBXPRESS}
var
Rep,FileName       : string;
DBODBC             : TDatabase;
tbSauver           : TTable;
QP                 : TQuery;
St                 : string;
{$ENDIF}
begin
  inherited;
{$IFNDEF DBXPRESS}
            Rep := ExtractFileDir (Application.ExeName);
            FileName := Rep+'\parametre\'+ DMImport.GzImpCorresp.TableName;

            DBODBC              := TDatabase.Create(nil);
            DBODBC.DriverName   := 'STANDARD';
            DBODBC.Params.Add('PATH='+Rep+'\parametre\');
            DBODBC.DatabaseName := 'REPRISE';
            DBODBC.LoginPrompt  := false;
            DBODBC.Open;
            QP := TQuery.Create(nil);
            QP.DatabaseName := DBODBC.DatabaseName;
            QP.SQL.Add ('Select * from '+DMImport.GzImpCorresp.TableName);
            QP.Open;
            tbSauver := DMImport.GzImpCorresp;

            if tbSauver.Active then tbSauver.Close;
            tbSauver.Open;

            While not (QP.Eof) do
            begin
                  with tbSauver do
                  begin
                       if FindKey([QP.FindField('Ident').asinteger,QP.FindField('Tablename').asstring,QP.FindField('Profile').asstring]) then
                       begin
                            St := Format('Le Correspondant %s existe déjà, confirmez-vous sa modification  ?', [QP.FindField('Profile').asstring]);
                            if (PGIAsk(St, 'Conception')<> mryes) then
                            begin
                                         Close; QP.free; DBODBC.Close; DBODBC.destroy;
                                         exit;
                            end;
                            Edit;
                       end
                       else Insert;
                       FieldByName('Ident').Asinteger      :=  QP.FieldByName('Ident').Asinteger;
                       FieldByName('Tablename').AsString   :=  QP.FindField('Tablename').asstring;
                       FieldByName('Profile').AsString     :=  QP.FindField('Profile').asstring;
                       FieldByName('Champcode').asstring   :=  QP.FindField('Champcode').asstring;
                       FieldByName('Codedepart').asstring  :=  QP.FindField('Codedepart').asstring;
                       FieldByName('Codearrive').asstring  :=  QP.FindField('Codearrive').asstring;
                       FieldByName('Champnature').asstring :=  QP.FindField('Champnature').asstring;
                       FieldByName('Nature').asstring      :=  QP.FindField('Nature').asstring;
                       FieldByName('Commentaire').asstring :=  QP.FindField('Commentaire').asstring;
                       FieldByName('Domaine').asstring     :=  QP.FindField('Domaine').asstring;
                       Post;
                  end;
                  QP.next;
             end;
             tbSauver.Close;
             QP.Close;
             QP.free;
             DBODBC.Close;
             DBODBC.destroy;
             BChercheClick(Sender);
{$ENDIF}
end;

procedure TFCorresp.Enliste1Click(Sender: TObject);
begin
  inherited;
    CorrespDlg := TCorrespDlg.Create(nil);
    CorrespDlg.Script := nil;
    CorrespDlg.Champ := nil;
   try
    (* récupération de la table à partir d'un fichier*)
	if CorrespDlg.ShowModal = mrOK then
            SaveListeCorresp(TRUE);
   finally
        CorrespDlg.Libere;
        CorrespDlg.Free;
   end;
     BChercheClick(Sender);
end;

procedure TFCorresp.Partable1Click(Sender: TObject);
var
i         : integer;
FCor      : PListeCorresp;
TCorr     : TADOTable;
begin
  inherited;
    TCorr := TADOTable.create(Application);
    TCorr.TableName := DMImport.GzImpCorresp.TableName;
    TCorr.Connection := DMImport.Db as TADOConnection;

    FCreatCorresp := TFCreatCorresp.Create(Application);
    Try
        FCreatCorresp.ShowModal ;
    Finally
        for i:=0 To  FCreatCorresp.ListeInfoCorresp.Count-1 do
        begin
             FCor :=FCreatCorresp.ListeInfoCorresp.Items[i];
             with TCorr do
             begin
                  if Active then Close;
                  Open;
                  if Locate('Ident;Tablename;Profile',VarArrayOf([FCor^.Ident, FCor^.Tablename, FCor^.Profile]), [loCaseInsensitive]) then
                  begin
                      if (PGIAsk('Le profil '+ FCor^.Profile + 'existe déjà, voulez-vous le modifier : ?',
                      'Conception')<> mryes) then  begin Close; continue; end;
                      Edit;
                  end
                  else Insert;
                  FieldByName('Ident').AsInteger          := FCor^.Ident;
                  FieldByName('Tablename').Asstring       := FCor^.Tablename;
                  FieldByName('Profile').Asstring         := FCor^.Profile;
                  FieldByName('Champcode').Asstring       := FCor^.Champcode;
                  FieldByName('Codedepart').Asstring      := FCor^.Codedepart;
                  FieldByName('Codearrive').Asstring      := FCor^.Codedearrive;
                  FieldByName('Champnature').Asstring     := FCor^.Champnature;
                  FieldByName('Nature').Asstring          := FCor^.Nature;
                  FieldByName('Commentaire').Asstring     := FCor^.Commentaire;
                  FieldByName('Domaine').Asstring         := RendDomaine(FCor^.Domaine);
                  Post;
                  Close;
             end;
        end;
        FCreatCorresp.freeListe;
        FCreatCorresp.Free ;
        TCorr.free;
     End;
     BChercheClick(Sender);
end;

procedure TFCorresp.SaveListeCorresp(Ask : Boolean);
var
TCorr      : TADOTable;
AStream    : TStream;
N          : Integer;
ScriptCorr : TScriptCorresp;
ABlobField : TBlobField;
Ident      : integer;
function MaxIdent : integer;
var
Q1 : TQuery;
begin
            Result := 1;
            Q1 := OpenSQLADO ('SELECT Max (Ident) Identification from '+DMImport.GzImpCorresp.TableName+
            ' Where Profile="'+ CorrespDlg.Profile.Text+'"', DMImport.Db);
            if not Q1.EOF then
            Result := Q1.findfield('Identification').asinteger+1;
            Q1.close; Q1.free;
end;
begin
            Ident := MaxIdent;
            TCorr := TADOTable.Create(Application);
            TCorr.Connection := DMImport.Db as TADOConnection;
            TCorr.TableName := DMImport.GzImpCorresp.TableName;

             with TCorr do
             begin
                  if Active then Close;
                  Open;
                 // if FindKey([Ident, TraduireMemoire('Liste de correspondance'), CorrespDlg.Profile.Text]) then
                 if Locate('Profile;Champcode',VarArrayOf([CorrespDlg.Profile.Text,CorrespDlg.Famille.Text]), [loCaseInsensitive]) then
                  begin
                       if Ask then
                          if (PGIAsk('Le profil '+ CorrespDlg.Profile.Text + 'existe déjà, voulez-vous le modifier : ?',
                             'Conception')<> mryes) then  begin Close; exit; end;
                      Edit;
                      Ident := FieldByName('Ident').asinteger;
                  end
                  else Insert;
                  FieldByName('Ident').AsInteger          := Ident;
                  FieldByName('Tablename').Asstring       := TraduireMemoire('Liste de correspondance');
                  FieldByName('Profile').Asstring         := CorrespDlg.Profile.Text;
                  FieldByName('Champcode').Asstring       := CorrespDlg.Famille.Text;
                  FieldByName('Codedepart').Asstring      := '';
                  FieldByName('Codearrive').Asstring      := '';
                  FieldByName('Champnature').Asstring     := '';
                  FieldByName('Nature').Asstring          := '';
                  FieldByName('Commentaire').Asstring     := CorrespDlg.Memo1.text;
                  FieldByName('Domaine').Asstring         := RendDomaine((CorrespDlg.Domaine.text));
                  ScriptCorr := TScriptCorresp.create (nil);
                  With ScriptCorr do
                  begin
                      Name       := CorrespDlg.Profile.Text;
                      FFichier    := CorrespDlg.edNomTable.Text;
                      for N:=0 to CorrespDlg.Table.FEntree.Count-1 do
                      begin
                           LFEntree.Add(CorrespDlg.Table.FEntree[N]);
                           LFSortie.add(CorrespDlg.Table.FSortie[N]);
                      end;
                  end;
                  ABlobField := FieldByName('TBLCOR') as TBlobField;
                  AStream := TBlobStream.Create(ABlobField, bmWrite);
                  ScriptCorr.SaveTo(AStream);
		              AStream.Free;
                  Post;
                  Close;  free;
                  ScriptCorr.Destroy;
             end;
end;

procedure TFCorresp.BinsertClick(Sender: TObject);
begin
  inherited;
    CorrespDlg := TCorrespDlg.Create(nil);
    CorrespDlg.Script := nil;
    CorrespDlg.Champ := nil;
   try
    (* récupération de la table à partir d'un fichier*)
	if CorrespDlg.ShowModal = mrOK then
            SaveListeCorresp(TRUE);
   finally
        CorrespDlg.Libere;
        CorrespDlg.Free;
   end;
     BChercheClick(Sender);
end;

procedure TFCorresp.BAnnulerClick(Sender: TObject);
begin
  inherited;
  Close;
  if IsInside(Self) then THPanel(parent).CloseInside;
end;

end.
