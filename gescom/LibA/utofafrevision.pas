{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 17/03/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFREVISION ()
Mots clefs ... : TOF;AFREVISION
*****************************************************************}
unit UtofAfRevision;   

interface

uses StdCtrls,
  Controls,
  Classes,
  {$IFDEF EAGLCLIENT}
  MainEagl,
  UtileAGL,
  {$ELSE}          
  Fe_Main,
  db,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  EdtEtat,
  {$ENDIF}

  {$IFDEF BTP}
  CalcOLEGenericBTP,
  {$ENDIF}

  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox, vierge,
  UTOF, Utob, dialogs, HTB97, uTofAfRevFormuleEdit, 
  EntGC,uRecupSQLModele,FactUtil, dicobtp, UtilRevFormule;
                                                         
type
  TOF_AFREVISION = class(TOF)
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
  private
    fTOBS1: TOB;
    fTOBS2: TOB;
    YenaAGauche, YenaADroite: boolean;
    Clause_where: string;
    fBoAffaire: Boolean;
    GS: THGrid;
    procedure IndiceDansGrid(NumFormule: Integer; TobIndice: TOB);
    procedure FormuleExpression(NumFormule: integer; TobIndice: TOB);
  end;

procedure AglLanceFicheAFREVISION(cle, Action: string);

implementation

procedure AglLanceFicheAFREVISION(cle, Action: string);
begin
  AglLanceFiche('AFF', 'AFREVISION', '', cle, Action);
end;

procedure TOF_AFREVISION.OnArgument(S: string);
var Critere, Champ, valeur: string;
  X: integer;
begin
  inherited;
  Critere := (Trim(ReadTokenSt(S)));
  while (Critere <> '') do
  begin
    if Critere <> '' then
    begin
      X := pos('=', Critere);
      if x <> 0 then
      begin
        Champ := copy(Critere, 1, X - 1);
        Valeur := Copy(Critere, X + 1, length(Critere) - X);
      end else
        Champ := Critere;
      if Champ = 'Clause_where' then Clause_where := Valeur;
      if Champ = 'AFFAIRE' then fBoAffaire := True;
      if Champ = 'LIGNEAFFAIRE' then fBoAffaire := False;
    end;
    Critere := (Trim(ReadTokenSt(S)));
  end;
end;

procedure TOF_AFREVISION.OnLoad;
var
  S1, S2, Sautre: string;
  QS1, QS2, Qautre: Tquery;
  TobAutre: Tob;
  Clause_where1, Clause_where2: string;
  vTOBS1: Tob;
  vTOBS2: Tob;
  i_ind : integer;
  i     : Integer;
begin

  inherited;
  setcontroltext('AFF_TIERS', '');
  setcontroltext('T_LIBELLE', '');
  setcontroltext('CODEAFFAIRE', '');
  setcontroltext('LIBAFFAIRE', '');
  setcontroltext('CODEARTICLE', '');
  setcontroltext('LIBARTICLE', '');

  setcontroltext('AFR_DERNIERCOEF1', '');
  setcontroltext('AFR_FORCODE1', '');
  setcontroltext('AFR_COEFCALC1', '');
  setcontroltext('AFR_COEFAPPLIQUE1', '');
  setcontroltext('AFR_DATECALCCOEF1', '');

  setcontroltext('AFR_DERNIERCOEF2', '');
  setcontroltext('AFR_FORCODE2', '');
  setcontroltext('AFR_COEFCALC2', '');
  setcontroltext('AFR_COEFAPPLIQUE2', '');
  setcontroltext('AFR_DATECALCCOEF2', '');

  for i_ind := 1 to 2 do
  begin

    GS := THGrid(GetControl('GS' + inttostr(i_ind)));

    GS.colwidths[0] := 100;
    GS.ColAligns[0] := taLeftJustify;

    for i := 1 to 10 do
    begin
      GS.ColAligns[i] := taRightJustify;
      GS.colwidths[i] := 65;
    end;

    // titres
    GS.Cells[0,1] := 'Indice';
    GS.Cells[0,2] := 'Valeur Initiale';
    GS.Cells[0,3] := 'Valeur Actualisée';
    GS.Cells[0,4] := 'Date Valeur';     
    GS.Cells[0,5] := 'Date Publication';
    GS.Cells[0,6] := 'Date Application';
    GS.Cells[0,7] := 'Journal';
    GS.Cells[0,8] := 'N° publication';
  end;

  if fBoAffaire then
    Clause_where := Clause_where + ' AND AFR_NUMERO=0 ';

  // remplissage des champs annexe dependant seulement de l'affaire
  Sautre := ' SELECT AFF_AFFAIRE,AFF_TIERS,AFF_LIBELLE,T_LIBELLE FROM AFREVISION ';
  Sautre := Sautre + ' INNER JOIN AFFAIRE ON AFREVISION.AFR_AFFAIRE = AFFAIRE.AFF_AFFAIRE ';
  Sautre := Sautre + 'LEFT JOIN TIERS ON AFFAIRE.AFF_TIERS = TIERS.T_TIERS ';
  Sautre := Sautre + Clause_where;
  QAutre := nil;
  TOBAutre := TOB.Create('', nil, -1);
  try
    QAutre := OpenSQL(Sautre, TRUE);
    TOBAutre.LoadDetailDB('', '', '', QAutre, false);
    if TOBAutre.Detail.count > 0 then
    begin
      setcontroltext('AFF_TIERS', TOBAutre.Detail[0].GetValue('AFF_TIERS'));
      setcontroltext('T_LIBELLE', TOBAutre.Detail[0].GetValue('T_LIBELLE'));
      setcontroltext('LIBAFFAIRE', TOBAutre.Detail[0].GetValue('AFF_LIBELLE'));
      {$IFDEF BTP}
      setcontroltext('CODEAFFAIRE', BTPCodeAffaireAffiche(TOBAutre.Detail[0].GetValue('AFF_AFFAIRE')));
      {$ELSE}
      setcontroltext('CODEAFFAIRE', CodeAffaireAffiche(TOBAutre.Detail[0].GetValue('AFF_AFFAIRE')));
      {$ENDIF}
      SetControlVisible('CODEARTICLE', false);
      SetControlVisible('LIBARTICLE', false);
      SetControlVisible('LBLIGNE', false);
    end;
  finally
    Ferme(QAutre);
    TOBAutre.free;
  end;

  // on va chercher les plus recents pour cote gauche et cote droit
  S1 := 'SELECT MAX(AFR_DATECALCCOEF) AS LADATE, AFR_AFFAIRE,AFR_FORCODE';
  S1 := S1 + ' FROM AFREVISION ';
  S1 := S1 + Clause_where;
  S1 := S1 + ' and AFR_NUMFORMULE =1 ';
  S1 := S1 + ' GROUP BY AFR_AFFAIRE,AFR_FORCODE';

  S2 := 'SELECT MAX(AFR_DATECALCCOEF) AS LADATE, AFR_AFFAIRE,AFR_FORCODE';
  S2 := S2 + ' FROM AFREVISION ';
  S2 := S2 + Clause_where;
  S2 := S2 + ' and AFR_NUMFORMULE =2 ';
  S2 := S2 + ' GROUP BY AFR_AFFAIRE,AFR_FORCODE';

  QS1 := nil;
  vTOBS1 := TOB.Create('', nil, -1);
  QS2 := nil;
  vTOBS2 := TOB.Create('', nil, -1);
  YenaAGauche := False;
  YenaADroite := False;
  try
    QS1 := OpenSQL(S1, TRUE);
    vTOBS1.LoadDetailDB('', '', '', QS1, false, true);
    if vTOBS1.Detail.count > 0 then
    begin
      YenaAGauche := True;
      Clause_where1 := ' WHERE (AFR_AFFAIRE = "' + vTOBS1.Detail[0].GetValue('AFR_AFFAIRE') + '" ';
      Clause_where1 := Clause_where1 + ' AND  AFR_DATECALCCOEF ="' + usdatetime(vTOBS1.Detail[0].GetValue('LADATE')) + '") ';
      Clause_where1 := Clause_where1 + ' AND  AFR_FORCODE = "' + vTOBS1.Detail[0].GetValue('AFR_FORCODE') + '" ';
      Clause_where1 := Clause_where1 + ' AND AFR_NUMFORMULE =1 ';
    end;
    QS2 := OpenSQL(S2, TRUE);
    vTOBS2.LoadDetailDB('', '', '', QS2, false, true);
    if vTOBS2.Detail.count > 0 then
    begin
      YenaADroite := True;
      Clause_where2 := ' WHERE (AFR_AFFAIRE = "' + vTOBS2.Detail[0].GetValue('AFR_AFFAIRE') + '" ';
      Clause_where2 := Clause_where2 + ' AND  AFR_FORCODE = "' + vTOBS2.Detail[0].GetValue('AFR_FORCODE') + '" ';
      Clause_where2 := Clause_where2 + ' AND  AFR_DATECALCCOEF =" ' + usdatetime(vTOBS2.Detail[0].GetValue('LADATE')) + '") ';
      Clause_where2 := Clause_where2 + ' AND AFR_NUMFORMULE =2 ';
    end;

  finally
    Ferme(QS1);
    vTOBS1.Free;
    Ferme(QS2);
    vTOBS2.Free;
  end;

  // on a fabriqué la clause where de chaque cote on va creer 2 tob et on prendra pour chaque cote le premier de chacune
  s1 := 'SELECT AFR_DATECALCCOEF, AFR_AFFAIRE, AFR_OKCOEFAPPLIQUE,  ';
  s1 := s1 + 'AFR_DERNIERCOEF,AFR_COEFAPPLIQUE,AFR_COEFCALC,AFR_NUMFORMULE,';
  for i_ind := 1 to 10 do
  begin
    s1 := s1 + 'AFE_INDCODE' + inttostr(i_ind) + ',AFR_VALINIT' + inttostr(i_ind) + ',AFR_VALINDICE' + inttostr(i_ind);
    s1 := s1 + ',AFC_INDAFF' + inttostr(i_ind) + ',AFR_DATEINDICE' + inttostr(i_ind);
    s1 := s1 + ',AFR_INDDATEVAL' + inttostr(i_ind);
    s1 := s1 + ',AFR_PUBCODE' + inttostr(i_ind) + ',';
  end;
  s1 := s1 + 'AFR_FORCODE,AFR_NUMEROLIGNE,AFE_FOREXPRESSION ';
  s1 := s1 + 'FROM AFREVISION ';
  s1 := s1 + 'LEFT JOIN AFFORMULE ON AFE_FORCODE=AFR_FORCODE ';
  s1 := s1 + 'LEFT JOIN AFPARAMFORMULE ON AFC_FORCODE=AFR_FORCODE AND AFC_AFFAIRE=AFR_AFFAIRE ';
  s1 := s1 + Clause_where1;

  QS1 := nil;
  fTOBS1 := TOB.Create('AFREVISION', nil, -1);
  if YenaAGauche then
  begin
    try
      QS1 := OpenSQL(s1, TRUE);
      fTOBS1.LoadDetailDB('', '', '', QS1, false, true);
      if fTOBS1.Detail.count > 0 then
      begin
        setcontroltext('AFR_DERNIERCOEF1', fTOBS1.Detail[0].GetValue('AFR_DERNIERCOEF'));
        setcontroltext('AFR_COEFAPPLIQUE1', fTOBS1.Detail[0].GetValue('AFR_COEFAPPLIQUE'));
        setcontroltext('AFR_COEFCALC1', fTOBS1.Detail[0].GetValue('AFR_COEFCALC'));
        setcontroltext('AFR_FORCODE1', fTOBS1.Detail[0].GetValue('AFR_FORCODE'));
        setcontroltext('AFR_COEFAPPLIQUEOK1', fTOBS1.Detail[0].GetValue('AFR_OKCOEFAPPLIQUE'));
        setcontroltext('AFR_DATECALCCOEF1', fTOBS1.Detail[0].GetValue('AFR_DATECALCCOEF'));
      end;
      IndiceDansGrid(1, fTOBS1.Detail[0]);
      FormuleExpression(1, fTOBS1.Detail[0]);
    finally
      Ferme(QS1);
    end;
  end
  else
  begin
    SetControlEnabled('AFR_COEFCALC1', false);
  end;

  s2 := 'SELECT AFR_DATECALCCOEF, AFR_AFFAIRE, AFR_OKCOEFAPPLIQUE, ';
  s2 := s2 + 'AFR_DERNIERCOEF,AFR_COEFAPPLIQUE,AFR_COEFCALC,AFR_NUMFORMULE,';
  for i_ind := 1 to 10 do
  begin
    s2 := s2 + 'AFE_INDCODE' + inttostr(i_ind) + ',AFR_VALINIT' + inttostr(i_ind) + ',AFR_VALINDICE' + inttostr(i_ind);
    s2 := s2 + ',AFC_INDAFF' + inttostr(i_ind) + ',AFR_DATEINDICE' + inttostr(i_ind);
    s2 := s2 + ',AFR_INDDATEVAL' + inttostr(i_ind);
    s2 := s2 + ',AFR_PUBCODE' + inttostr(i_ind) + ',';
  end;
  s2 := s2 + 'AFR_FORCODE,AFR_NUMEROLIGNE,AFE_FOREXPRESSION ';
  s2 := s2 + 'FROM AFREVISION ';
  s2 := s2 + 'LEFT JOIN AFFORMULE ON AFE_FORCODE=AFR_FORCODE ';
  s2 := s2 + 'LEFT JOIN AFPARAMFORMULE ON AFC_FORCODE=AFR_FORCODE AND AFC_AFFAIRE=AFR_AFFAIRE ';
  s2 := s2 + Clause_where2;

  QS2 := nil;
  fTOBS2 := TOB.Create('AFREVISION', nil, -1);
  if YenaADroite then
  begin
    try
      QS2 := OpenSQL(s2, TRUE);
      fTOBS2.LoadDetailDB('', '', '', QS2, false, true);
      if fTOBS2.Detail.count > 0 then
      begin
        setcontroltext('AFR_DERNIERCOEF2', fTOBS2.Detail[0].GetValue('AFR_DERNIERCOEF'));
        setcontroltext('AFR_COEFAPPLIQUE2', fTOBS2.Detail[0].GetValue('AFR_COEFAPPLIQUE'));
        setcontroltext('AFR_COEFCALC2', fTOBS2.Detail[0].GetValue('AFR_COEFCALC'));
        setcontroltext('AFR_FORCODE2', fTOBS2.Detail[0].GetValue('AFR_FORCODE'));
        setcontroltext('AFR_COEFAPPLIQUEOK2', fTOBS2.Detail[0].GetValue('AFR_OKCOEFAPPLIQUE'));
        setcontroltext('AFR_DATECALCCOEF2', fTOBS2.Detail[0].GetValue('AFR_DATECALCCOEF'));
      end;
      IndiceDansGrid(2, fTOBS2.Detail[0]);
      FormuleExpression(2, fTOBS2.Detail[0]);
    finally
      Ferme(QS2);
    end;
  end
  else
  begin
    SetControlEnabled('AFR_COEFCALC2', false);
  end;
  SetFocusControl('AFR_COEFCALC1');
end;

procedure TOF_AFREVISION.FormuleExpression(NumFormule: integer; TobIndice: TOB);
var StMemo: string;
    i, PosVirgule, PosParentheseFermee: integer;
    StaJeter, StIndice, StIndiceAff: string;
    vMemoDetail  : TMemo;
    vMemoEdition : TMemo;

begin
  StMemo := TobIndice.getvalue('AFE_FOREXPRESSION');
  while (pos(';', StMemo) > 0) do
  begin
    PosVirgule := pos(';', StMemo);
    PosParentheseFermee := pos('}', StMemo);
    StaJeter := copy(StMemo, PosVirgule, PosParentheseFermee - PosVirgule + 1);
    StMemo := Stringreplace(StMemo, StaJeter, '', [rfReplaceAll, rfIgnoreCase]);
  end;

  StMemo := FindEtReplace(StMemo,'ARR{','',true);
 // StMemo := FindEtReplace(StMemo,'[','',true);
 // StMemo := FindEtReplace(StMemo,']','',true);

  for i := 1 to 10 do
  begin
    StIndice := trim(TobIndice.getvalue('AFE_INDCODE' + inttostr(i)));
    StIndiceAff := trim(TobIndice.getvalue('AFC_INDAFF' + inttostr(i)));
    StMemo := Stringreplace(StMemo, StIndice, StIndiceAff, [rfReplaceAll, rfIgnoreCase]);
  end;

  vMemoDetail := TMemo.create(Ecran);
  vMemoDetail.Parent := Ecran;
  vMemoEdition := TMemo(GetControl('MEMO' + IntToStr(NumFormule)));

  try
    if not rien(StMemo) then
    begin
      RemplaceIndice(TobIndice.GetValue('AFR_FORCODE'),
                     TobIndice.GetValue('AFR_AFFAIRE'),
                     StMemo);     

      TraitementFormule(StMemo, vMemoDetail, vMemoEdition);
    end;
  vMemoEdition.Lines[0] := vMemoEdition.Lines[0];
  vMemoEdition.Lines[1] := vMemoEdition.Lines[1];
  vMemoEdition.Lines[2] := vMemoEdition.Lines[2];
  finally                         
    vMemoDetail.Free;
  end;
end;
  
procedure TOF_AFREVISION.IndiceDansGrid(NumFormule: Integer; TobIndice: TOB);
var
  i   : integer;
  vQr : Tquery;
  vSt : String;
//  vdt : tDatetime;
begin

  GS := THGrid(GetControl('GS' + inttostr(NumFormule)));
  for i := 1 to 10 do
  begin
    GS.Cells[i, 1] := TobIndice.getvalue('AFE_INDCODE' + inttostr(i));

    if TobIndice.getvalue('AFR_VALINDICE' + inttostr(i)) <> 0 then
      GS.Cells[i, 3] := TobIndice.getvalue('AFR_VALINDICE' + inttostr(i));

    if TobIndice.getvalue('AFR_INDDATEVAL' + inttostr(i)) > iDate1900 then
    begin
      vSt := 'SELECT AFV_INDNUMPUB, AFV_INDDATEPUB ';
      vSt := vSt + ' FROM AFVALINDICE ';
      vSt := vSt + ' WHERE AFV_INDCODE = "' + TobIndice.getvalue('AFE_INDCODE' + inttostr(i)) + '"';
      vSt := vSt + ' AND AFV_PUBCODE = "' + TobIndice.getvalue('AFR_PUBCODE' + inttostr(i)) + '"';
      vSt := vSt + ' AND AFV_INDDATEVAL = "' + UsDateTime(TobIndice.getvalue('AFR_INDDATEVAL' + inttostr(i))) + '"';

      vQr := nil;
      Try
        vQr := OpenSql(vSt, True);
         if vQr.FindField('AFV_INDDATEPUB').AsDateTime > iDate1900 then
        GS.Cells[i, 5] := vQr.FindField('AFV_INDDATEPUB').AsSTring;

      if TobIndice.getvalue('AFR_INDDATEVAL' + inttostr(i)) > iDate1900 then
        GS.Cells[i, 6] := DatetoStr(TobIndice.getvalue('AFR_INDDATEVAL' + inttostr(i)));
     
      GS.Cells[i, 7] := TobIndice.getvalue('AFR_PUBCODE' + inttostr(i));
      GS.Cells[i, 8] := vQr.FindField('AFV_INDNUMPUB').AsSTring;


      Finally
        Ferme(vQr);
      End;
    end;

    try
      if TobIndice.getvalue('AFR_VALINIT' + inttostr(i)) <> 0 then
        GS.Cells[i, 2] := TobIndice.getvalue('AFR_VALINIT' + inttostr(i));
    except
      PGIINFOAF('Recalculer le coefficient, les valeurs intiales ont été renseignées avec une ancienne méthode','');
      exit;
    end;

    if TobIndice.getvalue('AFR_DATEINDICE' + inttostr(i)) > iDate1900 then
      GS.Cells[i, 4] := DatetoStr(TobIndice.getvalue('AFR_DATEINDICE' + inttostr(i)));

  end;
end;

procedure TOF_AFREVISION.OnUpdate;
var vSt: string;
  vStDt: string;
begin
  inherited;
  if YenaAGauche then
  begin
    if fTOBS1.Detail.count > 0 then
    begin
      vStDt := GetControltext('AFR_DATECALCCOEF1');
      vSt := 'UPDATE AFREVISION SET AFR_COEFAPPLIQUE = ' + variantToSql(strtofloat(GetControltext('AFR_COEFAPPLIQUE1')));
      vSt := vSt + ' WHERE AFR_DATECALCCOEF = "' + usdatetime(strtodate(vStDt)) + '"';
      vSt := vSt + ' AND AFR_AFFAIRE = "' + fTOBS1.Detail[0].GetValue('AFR_AFFAIRE') + '"';
      vSt := vSt + ' AND AFR_FORCODE = "' + fTOBS1.Detail[0].GetValue('AFR_FORCODE') + '"';
      ExecuteSql(vSt);
    end;
  end;

  if YenaADroite then
  begin
    if fTOBS2.Detail.count > 0 then
    begin
      vStDt := GetControltext('AFR_DATECALCCOEF2');
      vSt := 'UPDATE AFREVISION SET AFR_COEFAPPLIQUE = ' + variantToSql(strtofloat(GetControltext('AFR_COEFAPPLIQUE2')));
      vSt := vSt + ' WHERE AFR_DATECALCCOEF = "' + usdatetime(strtodate(vStDt)) + '"';
      vSt := vSt + ' AND AFR_AFFAIRE = "' + fTOBS2.Detail[0].GetValue('AFR_AFFAIRE') + '"';
      vSt := vSt + ' AND AFR_FORCODE = "' + fTOBS2.Detail[0].GetValue('AFR_FORCODE') + '"';
      ExecuteSql(vSt);
    end;
  end;
end;

procedure TOF_AFREVISION.OnClose;
begin
  inherited;
  fTOBS1.Free;
  fTOBS2.Free;
end;

initialization
  registerclasses([TOF_AFREVISION]);
end.
