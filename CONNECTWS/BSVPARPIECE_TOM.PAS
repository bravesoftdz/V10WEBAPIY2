{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 30/08/2017
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : BSVPARPIECE (BSVPARPIECE)
Mots clefs ... : TOM;BSVPARPIECE
*****************************************************************}
Unit BSVPARPIECE_TOM ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fiche,
     FichList,
{$else}
     eFiche,
     eFichList,
{$ENDIF}
     HTB97, 
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOM, 
     UTob,
     UconnectBSV ;

Type
  TOM_BSVPARPIECE = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    procedure OnLoadAlerte               ; override ;
    procedure OnAfterCommit              ; override ;
    procedure OnAfterRollBack            ; override ;
    procedure OnAfterBeginTrans          ; override ;
  private
    NaturePiece : string;
    TOBPARAMBSV : TOB;
    TOBBSVPARPIECE : TOB;
    TYPEBSV : THDBValComboBox;
    SAIVALEUR : THEdit;
    BTPVALEUR : THValcomboBox;
    BSVVALEUR : THValcomboBox;
    DATEVALEUR : THEdit;
    procedure TypeBSVChange (Sender : TObject);
    procedure ConstitueTOBPARAMBSV(NaturePiece: string);
    procedure ChargeListeBSV(Code: string);
    procedure CHangeValeur (sender : TObject);
    procedure DeleteAndOut (Sender : Tobject);
    end ;

Implementation

procedure TOM_BSVPARPIECE.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_BSVPARPIECE.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_BSVPARPIECE.OnUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_BSVPARPIECE.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_BSVPARPIECE.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_BSVPARPIECE.OnLoadRecord ;
begin
  Inherited ;
  TypeBSVChange (Self);
end ;

procedure TOM_BSVPARPIECE.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_BSVPARPIECE.OnArgument ( S: String ) ;
var critere,ChampMul,ValMul,ST : string;
    X : Integer;
begin
  Inherited ;
  ST := S;                                                      
  repeat
    Critere := Trim(ReadTokenSt(ST));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        ChampMul := copy(Critere, 1, x - 1);
        ValMul := copy(Critere, x + 1, length(Critere));
        if ChampMul = 'PIECE' then NaturePiece:= ValMul;
      end;
    end;
  until Critere = '';
  //
  TOBPARAMBSV := TOB.Create ('LE PARAM GED',nil,-1);
  TOBBSVPARPIECE := TOB.Create ('LES PARPIECE BSV',nil,-1);
  GetParamsUploadBSV ( TOBPARAMBSV);
  if TOBPARAMBSV.Detail.Count > 0 then
  begin
    ConstitueTOBPARAMBSV (NaturePiece);
  end;
  TYPEBSV := THDBValComboBox (GetControl('BP3_TYPEBSV'));
  TYPEBSV.OnChange := TypeBSVChange;
  SAIVALEUR := THEdit (GetControl('SAIVALEUR'));  SAIVALEUR.OnChange := CHangeValeur;
  BTPVALEUR := THValcomboBox (GetControl('BTPVALEUR')); BTPVALEUR.OnChange := ChangeValeur;
  BSVVALEUR := THValcomboBox (GetControl('BSVVALEUR')); BSVVALEUR.OnChange := ChangeValeur;
  DATEVALEUR := THEdit(GetControl('DATEVALEUR'));   DATEVALEUR.OnChange := ChangeValeur;
  TToolbarButton97(GetCOntrol('BInsert')).visible := False;
  TToolbarButton97(GetCOntrol('Bdelete')).visible := False;
  TToolbarButton97(GetCOntrol('BDELTOUT')).onclick := DeleteAndOut;
end ;

procedure TOM_BSVPARPIECE.OnClose ;
begin
  TOBPARAMBSV.Free;
  TOBBSVPARPIECE.free;
  Inherited ;
end ;

procedure TOM_BSVPARPIECE.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_BSVPARPIECE.OnLoadAlerte;
begin
  Inherited ;
end ;

procedure TOM_BSVPARPIECE.OnAfterBeginTrans;
begin
  Inherited ;
end ;

procedure TOM_BSVPARPIECE.OnAfterCommit;
begin
  Inherited ;
end ;

procedure TOM_BSVPARPIECE.OnAfterRollBack;
begin
  Inherited ;
end ;


procedure TOM_BSVPARPIECE.ConstitueTOBPARAMBSV (NaturePiece : string);

  function TypeERP(CodeBSv : string) : string;
  begin
    if CodeBSv = 'text' then Result := '000'
    else if codeBsv = 'list' then result := '002'
    else CodeBSv := '003';
  end;

  procedure constitueTOBBSV;
  var TOBL : TOB;
      II : Integer;
      QQ : TQuery;
  begin
    TOBBSVPARPIECE.ClearDetail;                
    for II := 0 to TOBPARAMBSV.Detail.count -1 do
    begin
      TOBL := TOB.Create('BSVPARPIECE',TOBBSVPARPIECE,-1);
      TOBL.SetString('BP3_NATUREPIECEG',NaturePiece);
      TOBL.SetInteger('BP3_CODE',TOBPARAMBSV.detail[II].GetInteger('ID'));
      TOBL.SetString('BP3_LIBELLE',TOBPARAMBSV.detail[II].GetString('NAME'));
      TOBL.SetString('BP3_TYPEBSV', TypeERP(TOBPARAMBSV.detail[II].GetString('TYPE')));
      TOBL.SetString('BP3_VALEUR',TOBPARAMBSV.detail[II].GetString('DEFAULTVALUE'));
    end;
    QQ := OpenSQL('SELECT * FROM BSVPARPIECE WHERE BP3_NATUREPIECEG="'+NaturePiece+'"',True,-1,'',True);
    if not QQ.eof then
    begin
      QQ.First;
      repeat
        TOBL := TOBBSVPARPIECE.findfirst(['BP3_CODE'],[QQ.findfield('BP3_CODE').AsInteger],True);
        if TOBL <> nil then
        begin
          TOBL.SetString('BP3_TYPEBSV',QQ.findfield('BP3_TYPEBSV').AsString);
          TOBL.SetString('BP3_VALEUR',QQ.findfield('BP3_VALEUR').AsString);
        end;
        QQ.Next;
      until QQ.Eof;
    end;
    ferme (QQ);
  end;
  
begin
  constitueTOBBSV;
  BEGINTRANS;
  TRY
    ExecuteSQL('DELETE FROM BSVPARPIECE WHERE BP3_NATUREPIECEG="'+NaturePiece+'"');
    if TOBBSVPARPIECE.detail.count > 0 then
    begin
      TOBBSVPARPIECE.setallmodifie(True);
      TOBBSVPARPIECE.InsertDB(nil);
    end;
    COMMITTRANS;
  EXCEPT
    ROLLBACK;
  END;
  TOBBSVPARPIECE.ClearDetail;
end;

procedure TOM_BSVPARPIECE.ChargeListeBSV(Code : string);
var TT : TOB;
    II : Integer;
begin
  BSVVALEUR.Clear;
  TT := TOBPARAMBSV.findfirst(['ID'],[Code],true);
  if TT <> nil then
  begin
    For II := 0 to TT.Detail.count -1 do
    begin
      BSVVALEUR.AddItem(TT.detail[II].getstring('VALEUR'),nil);
    end;
  end;
end;


procedure TOM_BSVPARPIECE.TypeBSVChange(Sender: TObject);
var II : Integer;
begin
  if TYPEBSV.Value='002' then
  begin
    ChargeListeBSV(GetField('BP3_CODE'));
  end;
  SAIVALEUR.Visible := (TYPEBSV.Value='000');
  BTPVALEUR.Visible := (TYPEBSV.Value='001');
  BSVVALEUR.Visible := (TYPEBSV.Value='002');
  DATEVALEUR.Visible := (TYPEBSV.Value='003');
  if TYPEBSV.Value = '000' then
  begin
    SAIVALEUR.Text := GetField('BP3_VALEUR');
  end else if TYPEBSV.Value = '001' then
  begin
    II := -1;
    II := BTPVALEUR.Values.IndexOf(GetField('BP3_VALEUR'));
    if II >= 0 then BTPVALEUR.ItemIndex := II;
  end else if TYPEBSV.Value = '002' then
  begin
    II := -1;
    II := BSVVALEUR.Items.IndexOf(GetField('BP3_VALEUR'));
    if II >= 0 then BSVVALEUR.ItemIndex := II;
  end else
  begin
    DATEVALEUR.Text := GetField('BP3_VALEUR');
  end;
end;

procedure TOM_BSVPARPIECE.CHangeValeur(sender: TObject);
begin
  if not (DS.State in [dsInsert, dsEdit]) then DS.edit; // pour passer DS.state en mode dsEdit
  if TYPEBSV.Value = '000' then
  begin
    SetField('BP3_VALEUR',SAIVALEUR.Text);
  end else if TYPEBSV.Value = '001' then
  begin
    SetField('BP3_VALEUR',BTPVALEUR.value);
  end else if TYPEBSV.Value = '002' then
  begin
    SetField('BP3_VALEUR',BSVVALEUR.Text);
  end else
  begin
    SetField('BP3_VALEUR',DATEVALEUR.Text);
  end;
end;

procedure TOM_BSVPARPIECE.DeleteAndOut(Sender: Tobject);
begin
  if PgiAsk ('Attention : Vous allez supprimer la relation avec BSV pour cette pièce.#13#10 Confirmez-vous ?')<>MrYes then exit;
  ExecuteSQL('DELETE FROM BSVPARPIECE WHERE BP3_NATUREPIECEG="'+NaturePiece+'"');
  Ecran.Close;
end;

Initialization
  registerclasses ( [ TOM_BSVPARPIECE ] ) ;
end.
