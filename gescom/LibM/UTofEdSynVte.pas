unit UTofEdSynVte;

interface

uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
  HCtrls, HEnt1, HMsgBox, UTOF, UTOB, Spin, UtilSynVte,
  {$IFDEF EAGLCLIENT}
  eQRS1, utileAGL,
  {$ELSE}
  QRS1, db, dbTables, EdtEtat, Fiche,
  {$ENDIF}
  AglInit, EntGC, HQry, HStatus;

type
  TOF_EdSynVte = class(TOF)
    procedure OnLoad; override;
    procedure OnUpdate; override;
    procedure OnClose; override;
    procedure OnArgument(Argument: string); override;
  private
    Critere: string;
    Prixttc, Prixach, jsansvte: boolean;
    annee_dem, mois_dem: word;
  end;

implementation

procedure TOF_EdSynVte.OnArgument(Argument: string);
var stArgument: string;
  Spin: TSpinEdit;
begin
  inherited;
  stArgument := Argument;
  Critere := uppercase(Trim(ReadTokenSt(stArgument)));
  if critere = 'AD' then Ecran.Caption := TraduireMemoire('Synthèse annuelle');
  if critere = 'AA' then Ecran.Caption := TraduireMemoire('Synthèse annuelle comparatif N-1');
  if critere = 'MA' then Ecran.Caption := TraduireMemoire('Synthèse mensuelle comparatif N-1');
  if critere = 'MD' then Ecran.Caption := TraduireMemoire('Synthèse mensuelle');
  updatecaption(Ecran);
  Spin := TSpinEdit(TFQRS1(Ecran).FindComponent('EDTANNEE'));
  if Spin <> nil then Spin.value := strtoInt(FormatDateTime('yyyy', Date));
  Spin := TSpinEdit(TFQRS1(Ecran).FindComponent('EDTMOIS'));
  if Spin <> nil then Spin.value := strtoInt(FormatDateTime('mm', Date));
end;

procedure TOF_EdSynVte.OnLoad;
var F: TFQRS1;
  Edit3: TSpinEdit;
  Edit4: TSpinEdit;
  Edit5: TCheckBox;
  st_where: string;
begin
  inherited;
  F := TFQRS1(Ecran);
  SetControlText('XX_WHERE_USER', '');
  SetControlText('XX_WHERE', '');
  if (critere = 'AA') or (critere = 'MA') then
    SetControlText('XX_WHERE_USER', 'A.GZS_UTILISATEUR="' + V_PGI.USer + '"')
  else
    SetControlText('XX_WHERE_USER', 'GZS_UTILISATEUR="' + V_PGI.USer + '"');

  edit3 := TSpinEdit(TFQRS1(F).FindComponent('EDTANNEE'));
  if critere = 'AA' then
  begin
    st_where := 'A.GZS_ANNEE=' + inttostr(edit3.Value);
    st_where := st_where + ' and (A.GZS_QTETOT<>0 or B.GZS_QTETOT<>0)';
    SetControlText('XX_WHERE', st_where);
  end;
  if critere = 'AD' then
    SetControlText('XX_WHERE', 'GZS_ANNEE=' + inttostr(edit3.Value));
  if Ecran.Name = 'GCEDSYVTEM' then
  begin
    edit4 := TSpinEdit(TFQRS1(F).FindComponent('EDTMOIS'));
    edit5 := TCheckBox(TFQRS1(F).FindComponent('EDTJSVTE'));
    if critere = 'MA' then
    begin
      st_where := 'A.GZS_ANNEE=' + inttostr(edit3.Value) + ' and A.GZS_MOIS=' + inttostr(edit4.Value);
      if not edit5.Checked then
        // Ceci permet d'avoir toutes les jours de vente d'un mois donné de l'année N et/ou N-1
        st_where := st_where + ' and (A.GZS_QTETOT<>0 or B.GZS_QTETOT<>0)';
      SetControlText('XX_WHERE', st_where);
    end;
    if critere = 'MD' then
    begin
      st_where := 'GZS_ANNEE=' + inttostr(edit3.Value) + ' and GZS_MOIS=' + inttostr(edit4.Value);
      SetControlText('XX_WHERE', st_where);
    end;
  end;

end;

procedure TOF_EdSynVte.OnUpdate;
var F: TFQRS1;
  stWhere, RegimePrix: string;
  QEtiq: TQuery;
  TobTemp, TobSyn: TOB;
  Edit: THEdit;
  Edit2: THValcombobox;
  iBorne, nbjour: integer;
  ctrl, ctrlName, signe: string;
  LaDate: TDateTime;
  Jour: Word;
begin
  inherited;
  F := TFQRS1(Ecran);
  Prixttc := False;
  Prixach := False;
  jsansvte := False;
  if TcheckBox(TFQRS1(F).FindComponent('EDTTTC')).State = cbChecked then
    Prixttc := True;
  if TcheckBox(TFQRS1(F).FindComponent('EDTPRACH')).State = cbChecked then
    Prixach := True;

  ExecuteSQL('DELETE FROM GCTMPSYNVTE WHERE GZS_UTILISATEUR = "' + V_PGI.USer + '" AND GZS_TRAIT = "SYN"');

  ctrl := 'GL_ETABLISSEMENT';
  ctrlName := 'ETABLISSEMENT';
  signe := '=';
  Edit2 := THValComboBox(TFQRS1(F).FindComponent(ctrlName));
  if (Edit2 <> nil) and (Edit2.Value <> '') and (Edit2.Value <> TraduireMemoire('<<Tous>>')) then
  begin
    if stWhere <> '' then stWhere := stWhere + ' AND ';
    stWhere := stWhere + ctrl + signe + '"' + Edit2.Value + '"';
  end;

  ctrl := 'GL_DATEPIECE';

  if Ecran.Name = 'GCEDSYVTEM' then
  begin
    if TcheckBox(TFQRS1(F).FindComponent('EDTJSVTE')).State = cbChecked then
      jsansvte := True;
    Edit := THEdit(TFQRS1(F).FindComponent('DATEPIECE'));
    Ladate := StrToDate(Edit.Text);
    DecodeDate(LaDate, annee_dem, mois_dem, Jour);
    Jour := 1;
    nbjour := 30;
    if (mois_dem = 1) or (mois_dem = 3) or (mois_dem = 5)
      or (mois_dem = 7) or (mois_dem = 8) or (mois_dem = 10)
      or (mois_dem = 12) then nbjour := 31;
    if (mois_dem = 2) then
    begin
      if (isLeapYear(annee_dem) = True) then nbjour := 29
      else nbjour := 28;
    end;
    LaDate := Encodedate(Annee_dem, Mois_dem, Jour);
    if stWhere <> '' then stWhere := stWhere + ' AND ';
    stWhere := stWhere + '((' + ctrl + ' >= "' + USDateTime(Ladate) + '"';
    Jour := nbjour;
    LaDate := Encodedate(Annee_dem, Mois_dem, Jour);
    if stWhere <> '' then stWhere := stWhere + ' AND ';
    stWhere := stWhere + ctrl + ' <= "' + USDateTime(Ladate) + '" )';
    if (critere = 'MA') then
    begin
      Jour := 1;
      if (mois_dem = 2) then
      begin
        if (isLeapYear(annee_dem) = True) then nbjour := 29
        else nbjour := 28;
      end;
      LaDate := Encodedate(Annee_dem - 1, Mois_dem, Jour);
      if stWhere <> '' then stWhere := stWhere + ' OR (';
      stWhere := stWhere + ctrl + ' >= "' + USDateTime(Ladate) + '"';
      Jour := nbjour;
      LaDate := Encodedate(Annee_dem - 1, Mois_dem, Jour);
      if stWhere <> '' then stWhere := stWhere + ' AND ';
      stWhere := stWhere + ctrl + ' <= "' + USDateTime(Ladate) + '" )';
    end;
    stWhere := stWhere + ')';
  end;

  if Ecran.Name = 'GCEDSYVTEA' then
  begin
    for iBorne := 1 to 2 do
    begin
      if iBorne = 1
        then
        begin
        ctrlName := 'DATEPIECE_DEB';
        signe := '>=';
      end
      else
      begin
        ctrlName := 'DATEPIECE_FIN';
        signe := '<=';
      end;
      Edit := THEdit(TFQRS1(F).FindComponent(ctrlName));
      Ladate := StrToDate(Edit.Text);
      DecodeDate(LaDate, annee_dem, mois_dem, Jour);
      if (Edit <> nil) and (Edit.Text <> '01/01/1900') and (Edit.Text <> '31/12/2099') then
      begin
        if stWhere <> '' then stWhere := stWhere + ' AND ';
        stWhere := stWhere + ctrl + signe + '"' + USDateTime(StrToDate(Edit.Text)) + '"';
      end;
    end;
  end;

  QEtiq := OpenSQL('SELECT GL_ETABLISSEMENT,GL_NATUREPIECEG,GL_DATEPIECE,' +
    'GL_SOUCHE,GL_NUMERO,GL_INDICEG,GL_NUMLIGNE,GL_DEVISE,GL_QTEFACT,GL_TOTALHT,' +
    'GL_TOTALTTC,GL_PMAP,(GL_PUTTC*GL_QTEFACT)/GL_PRIXPOURQTE as BRUTTC, ' +
    '(GL_PUHT*GL_QTEFACT)/GL_PRIXPOURQTE as BRUTHT FROM LIGNE' +
    ' WHERE GL_NATUREPIECEG="FFO" AND GL_ARTICLE<> "" AND GL_TYPEARTICLE<> "FI" AND ' +
    stWhere + ' ORDER BY GL_ETABLISSEMENT,' +
    ' GL_DATEPIECE,GL_SOUCHE,GL_NUMERO,GL_INDICEG,GL_NUMLIGNE', True);

  if QEtiq.Eof then
  begin
    Ferme(QEtiq);
    exit;
  end;

  RegimePrix := TraduireMemoire('TTC') ;
  TobTemp := TOB.Create('', nil, -1);
  TobTemp.LoadDetailDB('', '', '', QEtiq, false);
  Ferme(QEtiq);
  TobSyn := TOB.Create('TMPSYNVTE', nil, -1);
  if ((critere = 'MA') or (jsansvte)) then RempliTousLesJours(TobTemp, Annee_dem, Mois_dem, TobSyn);
  if critere = 'AA' then RempliTousLesMois(TobTemp, Annee_dem, TobSyn);
  RempliTOBSynthese(TobTemp, TobSyn, critere, Prixttc, Prixach, 'SYN', '', '', 0, 0, critere);
  if TobSyn.Detail.Count > 0 then TobSyn.InsertDB(nil); //TobSyn.InsertOrUpdateDB();
  TobTemp.Free;
  TobSyn.Free;
end;

procedure TOF_EdSynVte.OnClose;
begin
  inherited;
  ExecuteSQL('DELETE FROM GCTMPSYNVTE WHERE GZS_UTILISATEUR = "' + V_PGI.USer + '" AND GZS_TRAIT = "SYN"');
end;

initialization
  registerclasses([TOF_EdSynVte]);

end.
