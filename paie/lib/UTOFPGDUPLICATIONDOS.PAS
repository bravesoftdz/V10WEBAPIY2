{***********UNITE*************************************************
Auteur  ...... : FC
Créé le ...... : 12/02/2007
Modifié le ... :   /  /
Description .. : Unit de duplication du paramètrage des éléments nationaux dossier
SUITE ........ : Duplication spécifique dans ce cas là car pas de notion de PREDEFINI et NODOSSIER
Mots clefs ... : PAIE;DUPLICATIONDOS
*****************************************************************
PT1 FC  14/06/2007 V_72 : Ne pas proposer le niveau Population si elles ne sont pas gérées
}
unit UTOFPGDUPLICATIONDOS;

interface
uses Classes, sysutils, HTB97, HCtrls, HMsgBox, UTOF, ParamDat, lookup
{$IFNDEF EAGLCLIENT}
{$IFNDEF DBXPRESS} ,dbtables {$ELSE} ,uDbxDataSet {$ENDIF}
{$ELSE}
,Utob
{$ENDIF}
;

type
  TOF_PGDUPLICATIONDOS = class(TOF)
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnClose; override;
    procedure Validation(Sender: TObject);
    procedure DateElipsisclick(Sender: TObject);
  private
    AncCode, AncNiveau, AncValNiveau: string;
    Longueur: integer;
    ParamTypeNiveau,StPop:String;
    Procedure ValNiveauElipsisClick(Sender : TObject);
    procedure MajFiltreNiveau(CodeElt : String);
  end;
implementation

uses Pgoutils, Hent1, PgOutils2;

{ TOF_PGDUPLICATIONDOS }
procedure TOF_PGDUPLICATIONDOS.DateElipsisclick(Sender: TObject);
var key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;

procedure TOF_PGDUPLICATIONDOS.OnArgument(Arguments: string);
var
  ECode, EDate: THEdit;
  Btn: TToolBarButton97;
  ValNiveau: THEdit;
  TypNiveau:THValComboBox;
begin
  DonneCodeDupliquerDos('', '', '', '');
  Longueur := 0;
  AncCode := ReadTokenst(Arguments);
  AncNiveau := ReadTokenst(Arguments);
  AncValNiveau := ReadTokenst(Arguments);
  Longueur := StrToInt(ReadTokenst(Arguments));
  ParamTypeNiveau := ReadTokenst(Arguments);

  SetControlProperty('CODEDUPLIQUER', 'MaxLength', Longueur);
  SetControlProperty('VALEURNIVEAU', 'value', '');
  SetControlProperty('DATEVALIDITE', 'text', Date);
  SetControlProperty('CODEDUPLIQUER', 'TEXT', AncCode);
  SetControlProperty('NIVEAU', 'TEXT', AncNiveau);

  //DEB PT1
{$IFNDEF CPS1}
  //S'il existe un parametre population
  if ExisteSQL('SELECT PPC_POPULATION FROM ORDREPOPULATION') then
    StPop := '"POP"'
  else
    StPop := '';
{$ENDIF}
  //FIN PT1

  if ParamTypeNiveau = 'SALARIE' then
  begin
    TypNiveau := THValComboBox(GetControl('NIVEAU'));
    TypNiveau.Value := 'SAL';
    TypNiveau.Enabled := False;
  end
  else if ParamTypeNiveau = 'ETAB' then
  begin
    TypNiveau := THValComboBox(GetControl('NIVEAU'));
    TypNiveau.Value := 'ETB';
    TypNiveau.Enabled := False;
  end
  else if ParamTypeNiveau = 'POP' then
  begin
    TypNiveau := THValComboBox(GetControl('NIVEAU'));
    TypNiveau.Value := 'POP';
    TypNiveau.Enabled := False;
  end
  else
    MajFiltreNiveau(AncCode);

  EDate := THEdit(GetControl('DATEVALIDITE'));
  if Edate <> nil then Edate.OnDblClick := DateElipsisclick;

  Btn := TToolBarButton97(GetControl('BValider'));
  if Btn <> nil then Btn.OnClick := Validation;

  ValNiveau := THEdit(GetControl('VALEURNIVEAU'));
  If ValNiveau <> Nil then
     ValNiveau.OnElipsisClick:= ValNiveauElipsisClick;
end;

Procedure TOF_PGDUPLICATIONDOS.ValNiveauElipsisClick(Sender : TObject);
var
  ValNiveau: THEdit;
  Niveau:THValComboBox;
begin
  ValNiveau := THEdit(GetControl('VALEURNIVEAU'));
  Niveau := THValComboBox(GetControl('NIVEAU'));
  if (Niveau<>nil) and (ValNiveau<>nil) then
  begin
    if Niveau.value = 'ETB' then
      ValNiveau.DataType := 'TTETABLISSEMENT';
    if Niveau.value = 'SAL' then
      ValNiveau.DataType := 'PGSALARIE';
    if Niveau.value = 'POP' then
      ValNiveau.DataType := 'PGPOPULATION';
    LookUpCombo(ValNiveau);
  end;
end;

procedure TOF_PGDUPLICATIONDOS.OnClose;
var
  ECode: THEdit;
begin
  inherited;
  Ecode := THEdit(GetControl('CODEDUPLIQUER'));
  if (ECode <> nil) then
  begin
    if Ecode.text = '' then
      DonneCodeDupliquerDos('', '', '', '');
  end;
end;

procedure TOF_PGDUPLICATIONDOS.Validation(Sender: TObject);
var
  ECode, EDate, EValNiveau: THEdit;
  ENiveau: THValComboBox;
  Btn: TToolBarButton97;
  OKRub, OkElt: Boolean;
  mes, CodeZero, temp, Niveau, ValNiveau, cc: string;
  Q : TQuery;
begin
  CodeZero := StringOfChar('0', Longueur);
  Btn := TToolBarButton97(GetControl('BValider'));
  Ecode := THEdit(GetControl('CODEDUPLIQUER'));
  ENiveau := THValComboBox(GetControl('NIVEAU'));
  EValNiveau := THEdit(GetControl('VALEURNIVEAU'));
  EDate := THEdit(GetControl('DATEVALIDITE'));
  if ENiveau <> nil then Niveau := ENiveau.value;
  if EValNiveau <> nil then ValNiveau := EValNiveau.text;
  if (ECode <> nil) and (ENiveau <> nil) and (EValNiveau <> nil) and (Btn <> nil) and (EDate <> nil) then
  begin
    DonneCodeDupliquerDos('', '', '', '');
    if (Niveau = '') and (ValNiveau <> '') and (EDate.text <> '') then
    begin
      PGIBox('Vous devez renseigner un niveau concerné!', 'Duplication');
      exit;
    end;
    if (Niveau <> '') and (ValNiveau = '') and (EDate.text <> '') then
    begin
      PGIBox('Vous devez renseigner une valeur pour le niveau concerné!', 'Duplication');
      exit;
    end;
    if (Niveau <> '') and (ValNiveau <> '') and (EDate.text = '') then
    begin
      PGIBox('Vous devez renseigner une date de validité!', 'Duplication');
      exit;
    end;

    // Vérifier que la valeur du niveau correspond au type du niveau : ETB, SAL ou POP
    if Niveau = 'ETB' then
      Q := OpenSQL('SELECT ET_ETABLISSEMENT FROM ETABLISS WHERE ET_ETABLISSEMENT="' + ValNiveau + '"', True);
    if Niveau = 'SAL' then
      Q := OpenSQL('SELECT PSA_SALARIE FROM SALARIES WHERE PSA_SALARIE="' + ValNiveau + '"', True);
    if Niveau = 'POP' then
      Q := OpenSQL('SELECT PPC_POPULATION FROM ORDREPOPULATION WHERE PPC_POPULATION="' + ValNiveau + '"', True);

    if (Niveau = 'ETB') or (Niveau = 'SAL') or (Niveau = 'POP') then
    begin
      if Q.Eof then
      begin
        PGIBox('La valeur niveau concerné ne correspond pas au niveau de l''élément national dossier!', 'Duplication');
        ferme(Q);
        exit;
      end;
      ferme(Q);
    end;

    if (Niveau <> '') and (ValNiveau <> '') and (EDate.text <> '') then
    begin
      PGIBox('Niveau et valeur corrects.#13#10#13#10Cliquer sur la croix pour terminer le traitement.', 'Duplication');
      DonneCodeDupliquerDos(Ecode.text, Niveau, ValNiveau, EDate.text);
      ENiveau.Enabled := False;
      EValNiveau.Enabled := False;
      EDate.Enabled := False;
      Btn.enabled := False;
    end;
  end; //Begin nil
end;

procedure TOF_PGDUPLICATIONDOS.MajFiltreNiveau(CodeElt : String);
var
  Q : TQuery;
  {$IFNDEF EAGLCLIENT}
  TypeNiveau : THDBValComboBox;
  {$ELSE}
  TypeNiveau : THValComboBox;
  {$ENDIF}
  NiveauOK, Plus : String;
begin
  if CodeElt <> '' then
  begin
    Q := OpenSQL('SELECT PNR_TYPENIVEAU,PNR_NIVMAXPERSO FROM ELTNIVEAUREQUIS WHERE PNR_CODEELT="' + CodeElt + '" AND ##PNR_PREDEFINI## ORDER BY PNR_PREDEFINI DESC', True);
    if not Q.Eof then
    begin
      {$IFNDEF EAGLCLIENT}
      TypeNiveau := THDBValComboBox(GetControl('NIVEAU'));
      {$ELSE}
      TypeNiveau := THValComboBox(GetControl('NIVEAU'));
      {$ENDIF}
      if TypeNiveau <> nil then
      begin
        NiveauOK := '""';
        if (Q.Fields[1].AsString <> 'CEG') and (Q.Fields[1].AsString <> 'STD') and (Q.Fields[1].AsString <> 'DOS') then
        begin
          if (Q.Fields[0].AsString = 'ETB') or(Q.Fields[0].AsString = 'CEG') or (Q.Fields[0].AsString = 'STD') or (Q.Fields[0].AsString = 'DOS') then
          begin
            NiveauOK := '"ETB"';
            if (Q.Fields[1].AsString = 'POP') and (StPop <> '') then
              NiveauOk := NiveauOK + ',"POP"';
            if (Q.Fields[1].AsString = 'SAL') then
            begin
              if StPop <> '' then
                NiveauOk := NiveauOK + ',"POP","SAL"'
              else
                NiveauOk := NiveauOK + ',"SAL"';
            end;
          end;
          if (Q.Fields[0].AsString = 'POP') then
          begin
            NiveauOK := StPop;
            if (Q.Fields[1].AsString = 'SAL') then
            begin
              if StPop <> '' then
                NiveauOk := NiveauOK + ',"SAL"'
              else
                NiveauOk := '"SAL"';
            end;
          end;
          if (Q.Fields[0].AsString = 'SAL') then
            NiveauOK := '"SAL"';
        end;
//        if NiveauOK <> '' then
          Plus := ' AND CO_CODE IN (' + NiveauOK + ')';

        TypeNiveau.Plus := Plus;
      end;
    end
    else
    begin
      //DEB PT1
      if StPop <> '' then
        StPop := StPop + ',';
      Plus := ' AND CO_CODE IN ("ETB",' + StPop + '"SAL")';
      {$IFNDEF EAGLCLIENT}
      TypeNiveau := THDBValComboBox(GetControl('NIVEAU'));
      {$ELSE}
      TypeNiveau := THValComboBox(GetControl('NIVEAU'));
      {$ENDIF}
      if TypeNiveau <> nil then
        TypeNiveau.Plus := Plus;
      //FIN PT1
    end;
    Ferme(Q);
  end;
end;

initialization
  registerclasses([TOF_PGDUPLICATIONDOS]);
end.
