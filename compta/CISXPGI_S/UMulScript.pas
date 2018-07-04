unit UMulScript;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls,
  DB, Hqry, Grids, DBGrids, StdCtrls, Hctrls,
  hmsgbox, ExtCtrls, HSysMenu, HDB, HTB97,
  HPanel, UiUtil, HEnt1, UPDomaine, HStatus, HRichEdt, ColMemo, ComCtrls,
  HRichOLE, Menus, Dialogs, PrintDbg, Mask, Mul,
{$IFDEF CISXPGI}
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet, ADODB,{$ENDIF}
{$ELSE}
  dbtables,
  {$IFDEF DBXPRESS} Variants, ADODB, {$ENDIF}
{$ENDIF}

  Forms;

function MulticritereScript (PP: THPanel; Var Dom : string) : String;

type
  TFMulScript = class(TFMul)
    HM: THMsgBox;
    QDOMAINE: TStringField;
    QCLE: TStringField;
    QCOMMENT: TStringField;
    QCLEPAR: TStringField;
    CLE: TEdit;
    XX_WHERE: TEdit;
    BDelete: TToolbarButton97;
    TDOMAINE: TLabel;
    QDATEDEMODIF: TDateField;
    Label2: TLabel;
    Label3: TLabel;
    CLEPAR: TEdit;
    Label4: TLabel;
    Comment: TEdit;
    Breprise: TToolbarButton97;
    BDelim: TToolbarButton97;
    QTable0: TStringField;
    QTable1: TStringField;
    Label1: TLabel;
    Label5: TLabel;
    Editeur: TEdit;
    Nature: TComboBox;
    ToolbarButton971: TToolbarButton97;
    HFindDialog1: THFindDialog;
    Domaine: THValComboBox;
    procedure BOuvrirClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure BinsertClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BChercheClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure VisibleClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure BrepriseClick(Sender: TObject);
    procedure BDelimClick(Sender: TObject);
    procedure DOMAINEChange(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
  private
    { Déclarations privées }
    CodeCle, DPar : string;
    ModeSelect    : Boolean;
    OkOuv         : Boolean;
    procedure SuppressionScript;
  public
    { Déclarations publiques }
  end;
var
  FMuls: TFMulScript;

implementation

{$R *.DFM}
uses UConcept, Uscript, UDMIMP, UAssistConcept;


function MulticritereScript (PP: THPanel; Var Dom : string) : String;
begin
 Result := '';
 FMuls := TFMulScript.Create(Application);
 FMuls.ModeSelect  := FALSE;
 FMuls.OkOuv       := FALSE;
 FMuls.DPar        := '';
 if (Dom <> '') and (Dom <> '0') then FMuls.DPar   := Dom;
 PP:=FindInsidePanel ;
 if PP=Nil then
 BEGIN
    FMuls.ModeSelect := TRUE;
    Try
     FMuls.ShowModal ;
    Finally
          if FMuls.OkOuv then
          begin
               Result := FMuls.Q.FindField ('CLE').asstring;
               Dom    := FMuls.Q.FindField ('DOMAINE').asstring;
          end;
     FMuls.Free ;
    End ;
 END else
 BEGIN
   InitInside(FMuls,PP) ;
   FMuls.Show ;
 END ;
end;

procedure TFMulScript.BOuvrirClick(Sender: TObject);
begin
  inherited;
  OkOuv := TRUE;
  FListeDblClick(Sender);
end;

procedure TFMulScript.FListeDblClick(Sender: TObject);
var
  Table, Dm, Nat: string;
  lib,Nt,edt    : string;
begin
  inherited;
  if ModeSelect then
  begin
       ModalResult := MrOk;
       OkOuv       := TRUE;
       exit;
  end;
  Table := Q.FindField('CLE').AsString;
  Dm    := Q.FindField('Domaine').AsString;
  Nat   := Q.FindField('CLEPAR').AsString;
  lib   := RendLibelleDomaine(Dm);
  Nt   := Q.FindField('Table0').AsString;
  Edt   := Q.FindField('Table1').AsString;
  PExtConcept(nil, Lib, Table, Nat, Nt, Edt, taModif);

end;

procedure TFMulScript.BinsertClick(Sender: TObject);
begin
  inherited;
    if (ParamCount > 0) and ((UpperCase(GetInfoVHCX.Mode) = 'EXPORT') or (UpperCase(GetInfoVHCX.Mode) = 'IMPORT')) then   exit;
  CreateScript (FALSE, tacreat);
  BChercheClick(Sender);
end;

procedure TFMulScript.FormCreate(Sender: TObject);
begin
  inherited;
  ChargementComboDomaine (Domaine);
  BImprimer.Enabled := FALSE;
end;

procedure TFMulScript.BChercheClick(Sender: TObject);
var
  SelectQL : string;
  SavePlace: TBookmark;
  St,DD    : string;
begin
//  SelectQL := ' SELECT Domaine,CLE,COMMENT,CLEPAR,DATEDEMODIF from '+ DMImport.GzImpReq.TableName;

  SelectQL := ' SELECT Table1,CLE,COMMENT,CLEPAR,DATEDEMODIF,Table0,Domaine from '+ DMImport.GzImpReq.TableName;
  Q.Close;
  Q.SQL.Clear;
  St := ' Where (CLEPAR<>"SQL" or CLEPAR="") ';
  if DPar <> '' then
  begin
     DD          := DPar;
     if Domaine.text = '' then
          Domaine.itemindex  := Domaine.items.IndexOf(RendLibelleDomaine(Dpar));
     if GetInfoVHCX.complement <> '' then Clepar.Text := GetInfoVHCX.complement;
     if GetInfoVHCX.Nature <> '' then
          Nature.Itemindex :=  Nature.items.IndexOf(GetInfoVHCX.Nature);
  end
  else
     DD :=  RendDomaine(Domaine.Text);
  if Cle.Text <> ''     then St := St + ' and CLE like "' + Cle.Text + '%" ';
  if (DD <> '') and (Domaine.Text <> '<Tous>') then St := St + ' and Domaine="' + DD + '" ';
  if Comment.Text <> '' then St := St + ' and Comment like "%' + Comment.Text + '%" ';
  if Clepar.Text <> ''  then St := St + ' and CLEPAR like "%' + Clepar.Text + '%" ';
  if (Nature.Text <> '') and (Nature.Text <> '<Tous>')  then St := St + ' and Table0 like "%' + Nature.Text + '%" ';
  if Editeur.Text <> '' then St := St + ' and Table1 like "' + Editeur.Text + '%" ';

  Q.SQL.Add(SelectQL + St + ' ORDER BY Table1,CLE');
  Q.ConnectionString := DMImport.DBGlobal.ConnectionString;
  //ChangeSQL(Q);
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


procedure TFMulScript.SuppressionScript;
var
{$IFDEF  DBXPRESS}
   tbcharger: TADOTable;
{$ELSE}
   tbcharger: TTable;
{$ENDIF}
begin
  {$IFDEF  DBXPRESS}
  tbcharger := DMImport.GzImpReq;
  {$ELSE}
  tbcharger := TTable.create(Application);
  tbcharger.DatabaseName := DMImport.DBGlobal.Name;
  {$ENDIF}
  tbcharger.Tablename := DMImport.GzImpReq.TableName;
  tbCharger.open;
  try
{$IFNDEF  DBXPRESS}
    if tbCharger.findKey([CodeCle]) then
{$ELSE}
    if tbCharger.Locate('CLE', CodeCle, [loCaseInsensitive]) then
{$ENDIF}
    begin
      tbCharger.Delete;
    end;
  except
    PGIBox(Exception(ExceptObject).message, '');
    exit;
  end;
  {$IFDEF  DBXPRESS}
  tbcharger.Close;
  {$ELSE}
  tbcharger.Free;
  {$ENDIF}


  {$IFDEF  DBXPRESS}
  tbcharger := DMImport.GzImpDelim;
  {$ELSE}
  tbcharger := TTable.create(Application);
  tbcharger.DatabaseName := DMImport.DBGlobal.Name;
  {$ENDIF}

  tbcharger.Tablename := DMImport.GzImpDelim.TableName;
  try
    tbCharger.open;
{$IFNDEF  DBXPRESS}
    if tbCharger.findKey([CodeCle]) then
{$ELSE}
    if tbCharger.Locate('CLE', CodeCle, [loCaseInsensitive]) then
{$ENDIF}
    begin
      tbCharger.Delete;
    end;
  except
  end;
  {$IFDEF  DBXPRESS}
  tbcharger.Close;
  {$ELSE}
  tbcharger.Free;
  {$ENDIF}

end;

procedure TFMulScript.BDeleteClick(Sender: TObject);
var
Texte : string;
i     : integer;
begin
  inherited;

  if (ParamCount > 0) and ((UpperCase(GetInfoVHCX.Mode) = 'EXPORT') or (UpperCase(GetInfoVHCX.Mode) = 'IMPORT')) then   exit;

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
      CodeCle := Q.FindField('Cle').AsString;
      SuppressionScript;
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
      CodeCle := Q.FindField('Cle').AsString;
      SuppressionScript;
    END;
    Fliste.ClearSelected;
  END;
  FiniMove;
  BChercheClick(Sender);

end;



procedure TFMulScript.VisibleClick(Sender: TObject);
begin
  inherited;
  BChercheClick(Sender);
end;

procedure TFMulScript.FormShow(Sender: TObject);
begin
  inherited;
  PComplement.TabVisible:=FALSE;
  if (Codeseria = '') or ((ParamCount > 0) and ((UpperCase(GetInfoVHCX.Mode) = 'EXPORT') or (UpperCase(GetInfoVHCX.Mode) = 'IMPORT'))) then
  begin
       Binsert.visible := FALSE;
       BDelete.visible := FALSE;
       BDelim.visible  := FALSE;
       BImprimer.visible := FALSE;
  end
  else
  begin
       Binsert.visible := TRUE;
       BDelete.visible := TRUE;
       BDelim.visible  := TRUE;
       BImprimer.visible := TRUE;
  end;
  if  (ParamCount > 0) and (GetInfoVHCX.Domaine <> '') then
  begin
          DOMAINEChange (Sender);
          Domaine.Enabled := FALSE;
          if GetInfoVHCX.complement <> '' then Clepar.Enabled := FALSE;
          if GetInfoVHCX.Nature <> '' then Nature.Enabled := FALSE;
  end;
end;

procedure TFMulScript.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
  case Key of
    VK_DELETE : BDeleteClick (Sender);
    VK_INSERT : BinsertClick (Sender);
  end;

end;

procedure TFMulScript.BrepriseClick(Sender: TObject);
{$IFNDEF DBXPRESS}
var
Rep,FileName       : string;
DBODBC             : TDatabase;
tbSauver           : TTable;
QP                 : TQuery;
St,CC              : string;
{$ENDIF}
begin
  inherited;
{$IFNDEF DBXPRESS}
            Rep := ExtractFileDir (Application.ExeName);
            FileName := Rep+'\parametre\'+ DMImport.GzImpReq.TableName;

            DBODBC              := TDatabase.Create(nil);
            DBODBC.DriverName   := 'STANDARD';
            DBODBC.Params.Add('PATH='+Rep+'\parametre\');
            DBODBC.DatabaseName := 'REPRISE';
            DBODBC.LoginPrompt  := false;
            DBODBC.Open;
            QP := TQuery.Create(nil);
            QP.DatabaseName := DBODBC.DatabaseName;
            QP.SQL.Add ('Select * from '+DMImport.GzImpReq.TableName+' Where (CLEPAR<>"SQL" or CLEPAR="") ');
            QP.Open;
            tbSauver := DMImport.GzImpReq;

            if tbSauver.Active then tbSauver.Close;
            tbSauver.Open;

            While not (QP.Eof) do
            begin
                  with tbSauver do
                  begin
                       CC := QP.FindField('Cle').asstring;
                       if FindKey([CC]) then
                       begin
                            St := Format('Le Script %s existe déjà, confirmez-vous sa modification  ?', [CC]);
                            if (PGIAsk(St, 'Conception')<> mryes) then
                            begin
                                 QP.next; continue;
                            end;
                            Edit;
                       end
                       else Insert;
                       FieldByName('PARAMETRES').AsVariant :=  QP.FieldByName('PARAMETRES').AsVariant;
                       FieldByName('TBLCOR').AsVariant     :=  QP.FieldByName('TBLCOR').AsVariant;
                       FieldByName('CLE').AsString         :=  CC;
                       FieldByName('COMMENT').AsString     :=  QP.FindField('Comment').asstring;
                       FieldByName('CLEPAR').AsString      :=  QP.FindField('CLEPAR').asstring;
                       FieldByName('MODIFIABLE').AsInteger :=  QP.FindField('MODIFIABLE').asinteger;
                       FieldByName('DOMAINE').asstring     :=  QP.FindField('DOMAINE').asstring;
                       FieldByName('DATEDEMODIF').asdatetime := now;
                       FieldByName('Table0').asstring := '';
                       FieldByName('Table1').asstring := '';
                       FieldByName('Table2').asstring := '';
                       FieldByName('Table3').asstring := '';
                       FieldByName('Table4').asstring := '';
                       FieldByName('Table5').asstring := '';
                       FieldByName('Table6').asstring := '';
                       FieldByName('Table7').asstring := '';
                       FieldByName('Table8').asstring := '';
                       FieldByName('Table9').asstring := '';

                       Post;
                  end;
                  QP.next;
             end;
             tbSauver.Close;
             QP.Close;
             QP.free;

            QP := TQuery.Create(nil);
            QP.DatabaseName := DBODBC.DatabaseName;
            QP.SQL.Add ('Select * from '+DMImport.GzImpDelim.TableName+' Where (CLEPAR<>"SQL" or CLEPAR="") ');
            QP.Open;
            tbSauver := DMImport.GzImpDelim;

            if tbSauver.Active then tbSauver.Close;
            tbSauver.Open;

            While not (QP.Eof) do
            begin
                  with tbSauver do
                  begin
                       CC := QP.FindField('Cle').asstring;
                       if FindKey([CC]) then
                       begin
                            St := Format('Le Script délimité %s existe déjà, confirmez-vous sa modification  ?', [CC]);
                            if (PGIAsk(St, 'Conception')<> mryes) then
                            begin
                                 QP.next; continue;
                            end;
                            Edit;
                       end
                       else Insert;
                       FieldByName('PARAMETRES').AsVariant :=  QP.FieldByName('PARAMETRES').AsVariant;
                       FieldByName('CLE').AsString         :=  CC;
                       FieldByName('COMMENT').AsString     :=  QP.FindField('Comment').asstring;
                       FieldByName('CLEPAR').AsString      :=  QP.FindField('CLEPAR').asstring;
                       FieldByName('MODIFIABLE').AsInteger :=  QP.FindField('MODIFIABLE').asinteger;
                       FieldByName('DOMAINE').asstring     :=  QP.FindField('DOMAINE').asstring;
                       FieldByName('DATEDEMODIF').asdatetime := now;
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

procedure TFMulScript.BDelimClick(Sender: TObject);
var
  Table, Nat    : string;
  Nt,edt,Dm     : string;
begin
  inherited;
  Table := Q.FindField('CLE').AsString;
  Nat   := Q.FindField('CLEPAR').AsString;
  Nt    := Q.FindField('Table0').AsString;
  Edt   := Q.FindField('Table1').AsString;
  Dm    := Q.FindField('Domaine').AsString;

  ModifScriptDelim(Table, Nat, Nt, Edt, Dm);
end;

procedure TFMulScript.DOMAINEChange(Sender: TObject);
begin
  inherited;
    if (RendDomaine(Domaine.Text) = 'O') then
    begin
      Nature.clear;
      Nature.Items.add ('<Tous>');
      Nature.Items.add ('Balance');
      Nature.Items.add ('Ecriture');
    end
    else
    begin
       Nature.clear;
       Nature.Items.add ('<Tous>');
       Nature.Items.add ('Dossier');
       Nature.Items.add ('Journal');
       Nature.Items.add ('Balance');
       Nature.Items.add ('Synchronisation');
    end;

end;

procedure TFMulScript.BAnnulerClick(Sender: TObject);
begin
  inherited;
  Close;
  if IsInside(Self) then THPanel(parent).CloseInside;
end;

end.

