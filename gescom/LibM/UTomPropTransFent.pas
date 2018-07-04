unit UTomPropTransFent;

interface
uses M3FP, StdCtrls, Controls, Classes, Dialogs, SaisUtil, HStatus, Grids,
  HCtrls, HEnt1, HMsgBox, Hdimension, UTOB, UTOM, AGLInit, EntGC,
  {$IFDEF EAGLCLIENT}
  Maineagl, eFichList,
  {$ELSE}
  dbTables, DBGrids, db, FE_Main, FichList,
  {$ENDIF}
  forms, sysutils, ComCtrls, HDB, Math, Messages, buttons, UtilArticle;

type
  TOM_PROPTRANSFENT = class(TOM)
  public
    FCreat: Boolean;
    procedure OnArgument(Arguments: string); override;
    procedure OnLoadRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnNewRecord; override;
    procedure ReinitProp;
    procedure TransfRef;
    procedure ConsultProp;
    procedure GenereTrans;
    procedure GeneCF;
    procedure EditProp;
    procedure NouvProp;
    procedure DuplicProp;
    procedure SetLastError(Num: integer; ou: string);
    procedure CheckEtat(NameControl: string);
  end;

const
  // libellés des messages
  TexteMessage: array[1..2] of string = (
    {1}'Vous devez saisir un libellé ',
    {2}'Vous devez saisir un dépôt'
    );

implementation

procedure TOM_PROPTRANSFENT.SetLastError(Num: integer; ou: string);
begin
  if ou <> '' then SetFocusControl(ou);
  LastError := Num;
  LastErrorMsg := TexteMessage[LastError];
end;

procedure TOM_PROPTRANSFENT.CheckEtat(NameControl: string);
begin
  SetControlChecked('VIDE', NameControl = 'VIDE');
  SetControlChecked('ENCOURS', NameControl = 'ENCOURS');
end;

procedure TOM_PROPTRANSFENT.OnArgument(Arguments: string);
var Critere, ChampMul, ValMul: string;
  x: integer;
begin
  inherited;
  FCreat := False;
  if Ecran.Name = 'PR_DUPLIC' then
  begin
    Ecran.Top := 315;
    Ecran.Left := 246;
    repeat
      Critere := uppercase(Trim(ReadTokenSt(Arguments)));
      if Critere <> '' then
      begin
        x := pos('=', Critere);
        if x <> 0 then
        begin
          ChampMul := copy(Critere, 1, x - 1);
          ValMul := copy(Critere, x + 1, length(Critere));
          if (ChampMul = 'PROPDEPART') or (ChampMul = 'LIBPTRF') then
          begin
            SetControlText(ChampMul, ValMul);
          end;
        end;
      end;
    until Critere = '';
  end;
end;

procedure TOM_PROPTRANSFENT.OnLoadRecord;
var Vide, EnCours: Boolean;
begin
  inherited;
  if Ecran.Name = 'PR_PROPTRF' then
  begin
    //Rafraichit la fiche
    Vide := False;
    EnCours := False;
    CheckEtat('');
    if (GetField('GTE_ENCOURSPTRF') = '-') and (GetField('GTE_GENEREPTRF') = '-')
      and (GetField('GTE_COMMANDEFOU') = '-') then
    begin
      CheckEtat('VIDE');
      Vide := True;
    end;
    if (GetField('GTE_ENCOURSPTRF') = 'X') then
    begin
      CheckEtat('ENCOURS');
      EnCours := True;
    end;

    TBitBtn(GetControl('BCONSULT')).Caption := 'Modification de la proposition';

    if GetField('GTE_CODEPTRF') = '' then
    begin
      CheckEtat('VIDE');
      FCreat := True;
      Vide := True;
      SetControlEnabled('BDUPLICATION', False);
    end
    else SetControlEnabled('BDUPLICATION', True);

    if Vide then
    begin
      SetControlEnabled('GTE_DEPOTEMET', True);
      SetControlEnabled('BREINIT', False);
      if FCreat = False then
      begin
        SetControlEnabled('BTRFREF', True);
        SetControlEnabled('BNOUVAFF', True);
      end else
      begin
        SetControlEnabled('BTRFREF', False);
        SetControlEnabled('BNOUVAFF', False);
      end;
      SetControlEnabled('BCONSULT', False);
      SetControlEnabled('BEDITION', False);
      SetControlEnabled('BGENERE', False);
      SetControlEnabled('BGENECF', False);
    end
    else
      if EnCours then
    begin
      SetControlEnabled('GTE_DEPOTEMET', False);
      SetControlEnabled('BREINIT', True);
      SetControlEnabled('BTRFREF', True);
      SetControlEnabled('BNOUVAFF', True);
      SetControlEnabled('BCONSULT', True);
      SetControlEnabled('BEDITION', True);
      SetControlEnabled('BGENERE', True);
      SetControlEnabled('BGENECF', True);
    end
    else
    begin
      SetControlEnabled('GTE_DEPOTEMET', False);
      SetControlEnabled('BREINIT', True);
      SetControlEnabled('BTRFREF', False);
      SetControlEnabled('BNOUVAFF', False);
      SetControlEnabled('BCONSULT', True);
      TBitBtn(GetControl('BCONSULT')).Caption := 'Consultation de la proposition';
      SetControlEnabled('BEDITION', True);
      SetControlEnabled('BGENECF', True);
      SetControlEnabled('BGENERE', True);
    end;

    //if GetField('GTE_COMMANDEFOU')='X' then SetControlEnabled('BGENECF',False);

    //if GetField('GTE_GENEREPTRF')='X' then SetControlEnabled('BGENERE',False);
  end;

  FCreat := False;
end;

procedure TOM_PROPTRANSFENT.OnUpdateRecord;
var TobProp, TobPropLig: TOB;
  TobProp2: TOB;
  i: integer;
  Num: string;
  SQL: string;
  Q: TQUERY;
begin
  inherited;
  if Ecran.Name = 'PR_DUPLIC' then
  begin
    TobProp := TOB.Create('PROPTRANSFENT', nil, -1);
    if TobProp.selectDB('"' + GetControlText('PROPDEPART') + '"', nil, False) then
    begin
      SetField('GTE_DEPOTEMET', TobProp.GetValue('GTE_DEPOTEMET'));
      if GetControlText('REPRISEPARAM') = 'X' then
      begin
        // Duplication du paramétrage de la proposition
        SetField('GTE_VIDDEPOT', TobProp.GetValue('GTE_VIDDEPOT'));
        SetField('GTE_STOCKMIN', TobProp.GetValue('GTE_STOCKMIN'));
        SetField('GTE_QTEFIXE', TobProp.GetValue('GTE_QTEFIXE'));
        SetField('GTE_TXSTOCK', TobProp.GetValue('GTE_TXSTOCK'));
        SetField('GTE_NBQTEFIXE', TobProp.GetValue('GTE_NBQTEFIXE'));
        SetField('GTE_NBTXSTOCK', TobProp.GetValue('GTE_NBTXSTOCK'));
        SetField('GTE_CONSULT', TobProp.GetValue('GTE_CONSULT'));
        SetField('GTE_REAFFECTART', TobProp.GetValue('GTE_REAFFECTART'));
        SetField('GTE_REAFFECTBTQ', TobProp.GetValue('GTE_REAFFECTBTQ'));
        SetField('GTE_ARRETRUPTURE', TobProp.GetValue('GTE_ARRETRUPTURE'));
        SetField('GTE_BLOCNOTE', TobProp.GetValue('GTE_BLOCNOTE'));
        for i := 1 to 10 do
        begin
          if (i < 10) then Num := IntToStr(i) else Num := 'A';
          SetField('GTE_METHODE' + Num, TobProp.GetValue('GTE_METHODE' + Num));
        end;
      end;
    end;
    TobProp.Free;
  end;

  if (GetField('GTE_LIBPTRF') = '') then
  begin
    SetLastError(1, 'GTE_LIBPTRF');
    exit;
  end;
  if (GetField('GTE_DEPOTEMET') = '') then
  begin
    SetLastError(2, 'GTE_DEPOTEMET');
    exit;
  end;
  if (getField('GTE_ENCOURSPTRF') = '-') and (getField('GTE_GENEREPTRF') = '-') then
  begin
    SetControlEnabled('BNOUVAFF', True);
    SetControlEnabled('BTRFREF', True);
  end;

  if Ecran.Name = 'PR_DUPLIC' then
  begin
    if GetControlText('REPRISEDONNEES') = 'X' then
    begin
      // Duplication des données de la proposition
      SQL := 'Select * from PROPTRANSFLIG where GTL_CODEPTRF="' + GetControlText('PROPDEPART') + '"';
      Q := OpenSQL(SQL, True);
      if not Q.EOF then
      begin
        SetField('GTE_ENCOURSPTRF', 'X');

        TobProp := TOB.Create('', nil, -1);
        TobProp.LoadDetailDB('PROPTRANSFLIG', '', '', Q, False, True);
        TobProp2 := TOB.Create('', nil, -1);
        TobProp2.Dupliquer(TobProp, True, True, True);
        for i := 0 to TobProp2.Detail.Count - 1 do
        begin
          TobPropLig := TobProp2.Detail[i];
          TobPropLig.PutValue('GTL_CODEPTRF', GetField('GTE_CODEPTRF'));
        end;
        TOBProp2.InsertDB(nil, True);
        TobProp.Free;
        TobProp2.Free;
      end;
      Ferme(Q);
    end;
  end;
end;

procedure TOM_PROPTRANSFENT.OnNewRecord;
begin
  inherited;
  FCreat := True;
  SetField('GTE_ENCOURSPTRF', '-');
  SetField('GTE_GENEREPTRF', '-');
  if Ecran.Name = 'PR_PROPTRF' then
  begin
    CheckEtat('VIDE');
    SetControlEnabled('BNOUVAFF', False);
    SetControlEnabled('BTRFREF', False);
    SetControlEnabled('BREINIT', False);
    SetControlEnabled('BCONSULT', False);
    SetControlEnabled('BEDITION', False);
    SetControlEnabled('BGENERE', False);
    SetControlEnabled('BGENECF', False);
  end;
  if Ecran.Name = 'PR_DUPLIC' then SetField('GTE_LIBPTRF', GetControlText('LIBPTRF'));
  SetField('GTE_STOCKMIN', 'X');
  SetField('GTE_VIDDEPOT', '-');
  SetField('GTE_QTEFIXE', '-');
  SetField('GTE_TXSTOCK', '-');
  SetField('GTE_NBQTEFIXE', 0);
  SetField('GTE_NBTXSTOCK', 0);
  SetField('GTE_CONSULT', '-');
  SetField('GTE_REAFFECTART', '-');
  SetField('GTE_REAFFECTBTQ', '-');
  SetField('GTE_ARRETRUPTURE', 'X');
end;

procedure TOM_PROPTRANSFENT.OnDeleteRecord;
var stReq: string;
begin
  inherited;
  stReq := 'delete from PROPTRANSFLIG where GTL_CODEPTRF="' + GetField('GTE_CODEPTRF') + '"';
  ExecuteSQL(stReq);
end;

// Duplication d'une proposition //

procedure TOM_PROPTRANSFENT.DuplicProp;
begin
  V_PGI.FormCenter := False;
  AGLLanceFiche('MBO', 'PR_DUPLIC', '', '', 'ACTION=CREATION;PROPDEPART=' + getField('GTE_CODEPTRF') + ';LIBPTRF=' + getField('GTE_LIBPTRF'));
  {$IFDEF EAGLCLIENT}
  {EAGLAFAIRE}
  {$ELSE}
  refreshDB;
  {$ENDIF}
end;

// Saisie d'une nouvelle affectation //

procedure TOM_PROPTRANSFENT.NouvProp;
begin
  DispatchArtMode(1, 'TYPETRAIT=PTR;CODETRAIT=' + getField('GTE_CODEPTRF'), '', 'GCARTICLE;PROPTRANS');
  {$IFNDEF EAGLCLIENT}
  Try
    refreshDB;
  Except
    //FMenuG.VireInside(nil);
    //TFFicheListe(Ecran).Close;
    SetControlEnabled('PAPPLI', False);
    Exit;
  end;
  {$ENDIF}
  if (getField('GTE_ENCOURSPTRF') = 'X') then
  begin
    SetControlEnabled('BREINIT', True);
    SetControlEnabled('GTE_DEPOTEMET', False);
    SetControlEnabled('BCONSULT', True);
    SetControlEnabled('BEDITION', True);
    SetControlEnabled('BGENERE', True);
    SetControlEnabled('BGENECF', True);
    CheckEtat('ENCOURS');
  end;
end;

// Saisie d'un transfert par référence //

procedure TOM_PROPTRANSFENT.TransfRef;
begin
  V_PGI.FormCenter := False;
  AGLLanceFiche('MBO', 'PR_TRANSFREF', '', '', GetField('GTE_CODEPTRF'));
  {$IFNDEF EAGLCLIENT}
  refreshDB;
  {$ENDIF}
  if (getField('GTE_ENCOURSPTRF') = 'X') then
  begin
    SetControlEnabled('BREINIT', True);
    SetControlEnabled('GTE_DEPOTEMET', False);
    SetControlEnabled('BCONSULT', True);
    SetControlEnabled('BEDITION', True);
    SetControlEnabled('BGENERE', True);
    SetControlEnabled('BGENECF', True);
    CheckEtat('ENCOURS');
  end;
end;

// Réinitialisation d'une propositions de transfert //

procedure TOM_PROPTRANSFENT.ReinitProp;
var stReq: string;
begin
  if PGIAsk('Confirmez vous la Réinitialisation de cette proposition de transfert ?', Ecran.Caption) <> mrYes then exit;
  stReq := 'delete from PROPTRANSFLIG where GTL_CODEPTRF="' + GetField('GTE_CODEPTRF') + '"';
  ExecuteSQL(stReq);
  stReq := 'Update PROPTRANSFENT set GTE_ENCOURSPTRF="-", GTE_GENEREPTRF="-", GTE_COMMANDEFOU="-" ' +
    'where GTE_CODEPTRF="' + GetField('GTE_CODEPTRF') + '"';
  ExecuteSQL(stReq);
  SetControlEnabled('BREINIT', False);
  SetControlEnabled('GTE_DEPOTEMET', False);
  SetControlEnabled('BCONSULT', False);
  SetControlEnabled('BEDITION', False);
  SetControlEnabled('BGENERE', False);
  SetControlEnabled('BGENECF', False);
  SetControlEnabled('BTRFREF', True);
  SetControlEnabled('BNOUVAFF', True);
  {$IFDEF EAGLCLIENT}
  {EAGLAFAIRE}
  {$ELSE}
  refreshDB;
  {$ENDIF}
  SetField('GTE_ENCOURSPTRF', '-');
  SetField('GTE_GENEREPTRF', '-');
  SetField('GTE_COMMANDEFOU', '-');
  CheckEtat('VIDE');
end;

// Consultation d'une proposition de transfert //

procedure TOM_PROPTRANSFENT.ConsultProp;
var TOBPropTransfLig: TOB;
  CodePtrf: string;
  Q: TQuery;
  PropGenere: boolean;
begin
  CodePtrf := GetControlText('GTE_CODEPTRF');
  Q := OpenSQL('SELECT GTE_ENCOURSPTRF,GTE_GENEREPTRF FROM PROPTRANSFENT where GTE_CODEPTRF="' + CodePtrf + '"', True);
  if Q.Eof then exit;
  PropGenere := Boolean(Q.Findfield('GTE_GENEREPTRF').AsString = 'X');
  Ferme(Q);
  if PropGenere then
    AGLLanceFiche('MBO', 'PR_AFFTRANSFERT', '', '', 'ACTION=CONSULTATION;CODEPTRF=' + CodePtrf)
  else
  begin
    AGLLanceFiche('MBO', 'PR_AFFTRANSFERT', '', '', 'ACTION=MODIFICATION;CODEPTRF=' + CodePtrf);
    TOBPropTransfLig := TheTOB;
    TheTOB := nil;
    // Enregistrement de la proposition d'affectation du transfert dans la Table
    if (TOBPropTransfLig <> nil) and (TOBPropTransfLig.Detail.Count > 0) then
      TOBPropTransfLig.UpdateDB(True);
    TOBPropTransfLig.Free;
  end;
end;

// Génération des transferts

procedure TOM_PROPTRANSFENT.GenereTrans;
var CodePtrf: string;
begin
  CodePtrf := GetControlText('GTE_CODEPTRF');
  //Appel de la fiche PR_GENETRANSF pour lancer la génération des transferts
  AGLLanceFiche('MBO', 'PR_GENETRANSF', '', '', CodePtrf);
  //Rafraichit la fiche
  refreshDB;
  OnLoadRecord;
end;

procedure TOM_PROPTRANSFENT.GeneCF;
var CodePtrf: string;
begin
  CodePtrf := GetControlText('GTE_CODEPTRF');
  //Appel de la fiche PR_GENECF pour lancer la génération des commandes fournisseurs
  AGLLanceFiche('MBO', 'PR_GENECF', '', '', CodePtrf);
  //Rafraichit la fiche
  refreshDB;
  OnLoadRecord;
end;

// Edition d'une proposition des transfert //

procedure TOM_PROPTRANSFENT.EditProp;
var CodePtrf: string;
begin
  CodePtrf := GetControlText('GTE_CODEPTRF');
  //Appel de la fiche lanceur
  AGLLanceFiche('MBO', 'CHOIXEDITTRANS', '', '', CodePtrf);
end;

// AGL /////////////////////////////////////////////////

procedure AGLOnNouvProp(Parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFicheListe)
    then OM := TFFicheListe(F).OM
  else exit;
  if (OM is TOM_PROPTRANSFENT) then TOM_PROPTRANSFENT(OM).NouvProp else exit;
end;

procedure AGLOnDuplicProp(Parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFicheListe)
    then OM := TFFicheListe(F).OM
  else exit;
  if (OM is TOM_PROPTRANSFENT) then TOM_PROPTRANSFENT(OM).DuplicProp else exit;
end;

procedure AGLOnReinitProp(Parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFicheListe)
    then OM := TFFicheListe(F).OM
  else exit;
  if (OM is TOM_PROPTRANSFENT) then TOM_PROPTRANSFENT(OM).ReinitProp else exit;
end;

procedure AGLOnTransfRef(Parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFicheListe)
    then OM := TFFicheListe(F).OM
  else exit;
  if (OM is TOM_PROPTRANSFENT) then TOM_PROPTRANSFENT(OM).TransfRef else exit;
end;

procedure AGLOnConsultProp(Parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFicheListe)
    then OM := TFFicheListe(F).OM
  else exit;
  if (OM is TOM_PROPTRANSFENT) then TOM_PROPTRANSFENT(OM).ConsultProp else exit;
end;

procedure AGLOnGenereTrans(Parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFicheListe) then OM := TFFicheListe(F).OM else exit;
  if (OM is TOM_PROPTRANSFENT) then TOM_PROPTRANSFENT(OM).GenereTrans else exit;
end;

procedure AGLOnGeneCF(Parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFicheListe) then OM := TFFicheListe(F).OM else exit;
  if (OM is TOM_PROPTRANSFENT) then TOM_PROPTRANSFENT(OM).GeneCF else exit;
end;

procedure AGLOnEditProp(Parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFicheListe)
    then OM := TFFicheListe(F).OM
  else exit;
  if (OM is TOM_PROPTRANSFENT) then TOM_PROPTRANSFENT(OM).EditProp else exit;
end;

initialization
  registerclasses([TOM_PROPTRANSFENT]);
  RegisterAglProc('OnReinitProp', True, 1, AGLOnReinitProp);
  RegisterAglProc('OnGenereTrans', True, 1, AGLOnGenereTrans);
  RegisterAglProc('OnGeneCF', True, 1, AGLOnGeneCF);
  RegisterAglProc('OnTransfRef', True, 1, AGLOnTransfRef);
  RegisterAglProc('OnConsultProp', True, 1, AGLOnConsultProp);
  RegisterAglProc('OnEditProp', True, 1, AGLOnEditProp);
  RegisterAglProc('OnNouvProp', True, 1, AGLOnNouvProp);
  RegisterAglProc('OnDuplicProp', True, 1, AGLOnDuplicProp);
end.
