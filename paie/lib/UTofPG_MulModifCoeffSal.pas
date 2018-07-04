{***********UNITE*************************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 11/09/2002
Modifié le ... :   /  /
Description .. : Unité de gestion du multicritère MODIFCOEFFSAL
Mots clefs ... : PAIE,PGMODIFCOEFFSAL
*****************************************************************}
{
PT1   : 15/10/2003 SB V_42 Maj de date modification de la table des salaries
                      pour introduction PAIE
PT2   : 11/05/2004 VG V_50 Affichage des éléments saisissables en fonction du
                           paramétrage stocké dans la table EVOLUTIONSAL
}
unit UTofPG_MulModifCoeffSal;

interface
uses UTOF, HCtrls, Hent1, hmsgbox, classes, UTOB, ed_tools, HStatus, EntPaie, PgOutils2,
  PgOutils, sysutils, HTB97, hqry, P5Def,
{$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} mul, HDB;
{$ELSE}
  EMul;
{$ENDIF}


type
  TOF_PGMULMODIFCOEFFSAL = class(TOF)
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;

  private
    BCherche: TToolbarButton97;
    WW: THEdit; // Clause XX_WHERE
    Q_Mul: THQuery; // Query pour changer la liste associee
    Amodifier: string;

    procedure ActiveWhere(Sender: TObject);
    procedure Reaffectation(Sender: TObject);
    procedure MAJSal(Champ: string);
  end;

implementation

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 11/09/2002
Modifié le ... :   /  /
Description .. : OnArgument
Mots clefs ... : PAIE;PGMODIFCOEFFSAL
*****************************************************************}

procedure TOF_PGMULMODIFCOEFFSAL.OnArgument(Arguments: string);
var
  St, StDelete: string;
  Q: TQuery;
  MajCoeff, MaxFin, NewMajCoeff: TDateTime;
  TModifCoeff, TModifCoeffFille: TOB;
  Num: Integer;
begin
  inherited;
  MaxFin := IDate1900;
  St := 'SELECT MAX(PPU_DATEFIN) AS MAXFIN FROM PAIEENCOURS';
  Q := OpenSql(St, True);
  if not Q.EOF then
  try
    MaxFin := Q.FindField('MAXFIN').AsDateTime;
  except
    on E: EConvertError do
      MaxFin := IDate1900;
  end;
  Ferme(Q);

  SetControlText('DATE', DateToStr(MaxFin));

  for Num := 1 to VH_Paie.PGNbreStatOrg do
  begin
    if Num > 4 then
      Break;
    VisibiliteChampSalarie(IntToStr(Num), GetControl('PSA_TRAVAILN' + IntToStr(Num)), GetControl('TPSA_TRAVAILN' + IntToStr(Num)));
  end;
  VisibiliteStat(GetControl('PSA_CODESTAT'), GetControl('TPSA_CODESTAT')); //PT3

  St := 'SELECT PMC_DATEANCIENFIN FROM MODIFCOEFF WHERE PMC_MODIFCOEFF = 0';
  Q := OpenSql(St, True);
  if (not Q.EOF) then
    MajCoeff := Q.FindField('PMC_DATEANCIENFIN').AsDateTime
  else
    MajCoeff := IDate1900; // PORTAGECWAS
  Ferme(Q);

  St := 'SELECT PMC_DATEANCIEN' +
    ' FROM MODIFCOEFF WHERE' +
    ' PMC_DATEANCIEN="' + UsDateTime(IDate1900) + '"';
  if ((MaxFin <> MajCoeff) or (ExisteSQL(St))) then
  begin
    St := 'SELECT * FROM MODIFCOEFF';
    Q := OpenSql(St, True);
    TModifCoeff := TOB.Create('Mère ModifCoeff', nil, -1);
    TModifCoeff.LoadDetailDB('MODIFCOEFF', '', '', Q, False);
    Ferme(Q);
    try
      begintrans;
      StDelete := 'DELETE FROM MODIFCOEFF';
      ExecuteSQL(StDelete);
      TModifCoeffFille := TModifCoeff.FindFirst([''], [''], FALSE);
      if (TModifCoeffFille <> nil) then
      begin
        while (TModifCoeffFille <> nil) do
        begin
          NewMajCoeff := PlusMois(MaxFin, -(TModifCoeffFille.GetValue('PMC_NBANCIENSUP'))) + 1;
          TModifCoeffFille.PutValue('PMC_DATEANCIEN', NewMajCoeff);
          NewMajCoeff := PlusMois(MaxFin, -(TModifCoeffFille.GetValue('PMC_NBANCIEN')));
          TModifCoeffFille.PutValue('PMC_DATEANCIENFIN', NewMajCoeff);
          TModifCoeffFille := TModifCoeff.FindNext([''], [''], FALSE);
        end;
      end
      else
      begin
        TModifCoeffFille := TOB.Create('MODIFCOEFF', TModifCoeff, -1);
        TModifCoeffFille.AddChampSup('PMC_MODIFCOEFF', FALSE);
        TModifCoeffFille.AddChampSup('PMC_NBANCIEN', FALSE);
        TModifCoeffFille.AddChampSup('PMC_NBANCIENSUP', FALSE);
        TModifCoeffFille.AddChampSup('PMC_DATEANCIEN', FALSE);
        TModifCoeffFille.AddChampSup('PMC_DATEANCIENFIN', FALSE);
        TModifCoeffFille.PutValue('PMC_MODIFCOEFF', 0);
        TModifCoeffFille.PutValue('PMC_NBANCIEN', 0);
        TModifCoeffFille.PutValue('PMC_NBANCIENSUP', 0);
        NewMajCoeff := PlusMois(MaxFin, 0);
        TModifCoeffFille.PutValue('PMC_DATEANCIEN', NewMajCoeff);
        TModifCoeffFille.PutValue('PMC_DATEANCIENFIN', NewMajCoeff);
      end;
      TModifCoeff.SetAllModifie(TRUE);
      TModifCoeff.InsertDB(nil, FALSE);
      FreeAndNil(TModifCoeff);

      CommitTrans;
    except
      Rollback;
      PGIBox('Une erreur est survenue lors de la mise à jour de la base', 'Coefficient');
    end;
  end;

  TFMul(Ecran).BOuvrir.OnClick := Reaffectation;
  Q_Mul := THQuery(Ecran.FindComponent('Q'));
  WW := THEdit(GetControl('XX_WHERE'));
  BCherche := TToolbarButton97(GetControl('BCherche'));
  ActiveWhere(nil);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 11/09/2002
Modifié le ... :   /  /
Description .. : OnLoad
Mots clefs ... : PAIE;PGMODIFCOEFFSAL
*****************************************************************}

procedure TOF_PGMULMODIFCOEFFSAL.OnLoad;
begin
  inherited;
  ActiveWhere(nil);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 11/09/2002
Modifié le ... :   /  /
Description .. : Réaffectation des coefficient des salariés
Mots clefs ... : PAIE;PGMODIFCOEFFSAL
*****************************************************************}

procedure TOF_PGMULMODIFCOEFFSAL.Reaffectation(Sender: TObject);
var
  champ: string;
  i: integer;
{$IFNDEF EAGLCLIENT}
  Liste: THDBGrid;
{$ELSE}
  Liste: THGrid;
{$ENDIF}
begin
{$IFNDEF EAGLCLIENT}
  Liste := THDBGrid(GetControl('FListe'));
{$ELSE}
  Liste := THGrid(GetControl('FListe'));
{$ENDIF}
  if Liste <> nil then
  begin
    if (Liste.NbSelected = 0) and (not Liste.AllSelected) then
    begin
      MessageAlerte('Aucun élément sélectionné');
      exit;
    end;

    if (Amodifier = '001') then
      champ := 'COEFFICIENT'
    else
      if (Amodifier = '002') then
        champ := 'QUALIFICATION'
      else
        if (Amodifier = '003') then
          champ := 'NIVEAU'
        else
          if (Amodifier = '004') then
            champ := 'INDICE';

    if (Liste.AllSelected = TRUE) then
    begin
      InitMoveProgressForm(nil, 'Réaffectation en cours', 'Veuillez patienter SVP ...', TFmul(Ecran).Q.RecordCount, FALSE, TRUE);
      InitMove(TFmul(Ecran).Q.RecordCount, '');

      TFmul(Ecran).Q.First;
      while not TFmul(Ecran).Q.EOF do
      begin
        MAJSal(Champ);
        TFmul(Ecran).Q.Next;
      end;

      Liste.AllSelected := False;
      TFMul(Ecran).bSelectAll.Down := Liste.AllSelected;
    end
    else
    begin
      InitMoveProgressForm(nil, 'Réaffectation en cours', 'Veuillez patienter SVP ...', Liste.NbSelected, FALSE, TRUE);
      InitMove(Liste.NbSelected, '');

      for i := 0 to Liste.NbSelected - 1 do
      begin
        Liste.GotoLeBOOKMARK(i);
        MAJSal(Champ);
      end;

      Liste.ClearSelected;
    end;

    FiniMove;
    FiniMoveProgressForm; // PORTAGECWAS
    if BCherche <> nil then
      BCherche.click;
    PGIBox('Traitement terminé', 'Réaffectation des coefficients des salariés');
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 11/05/2004
Modifié le ... :   /  /
Description .. : Mise à jour des champs salariés
Mots clefs ... : PAIE;PGMODIFCOEFFSAL
*****************************************************************}

procedure TOF_PGMULMODIFCOEFFSAL.MAJSal(Champ: string);
var
  Valeur, St: string;
begin
  MoveCur(False);
  St := TFmul(Ecran).Q.FindField('PSA_SALARIE').asstring;
  Valeur := TFmul(Ecran).Q.FindField('PMC_' + Champ).asstring;
  if St <> '' then
  begin
    if (isnumeric(St) and (VH_PAIE.PgTypeNumSal = 'NUM')) then
      St := ColleZeroDevant(StrToInt(St), 10);
    try
      begintrans;
      ExecuteSQL('UPDATE SALARIES SET PSA_' + Champ + ' = "' + Valeur + '",' +
        ' PSA_DATEMODIF="' + UsTime(Now) + '" ' + //PT1
        'WHERE PSA_SALARIE = "' + St + '"');
      CommitTrans;
    except
      Rollback;
      PGIBox('Une erreur est survenue lors de la mise à jour de la base',
        'Réaffectation des coefficients des salariés');
    end;
  end;
  MoveCurProgressForm(St); // PORTAGECWAS
end;


{***********A.G.L.***********************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 11/09/2002
Modifié le ... :   /  /
Description .. : ActiveWhere
Mots clefs ... : PAIE;PGMODIFCOEFFSAL
*****************************************************************}

procedure TOF_PGMULMODIFCOEFFSAL.ActiveWhere(Sender: TObject);
var
  Coeff, Indice, LibEmp, Niveau, Qualif, StDif, StEvol: string;
  QRechEvol: TQuery;
  TEvol: TOB;
begin
//PT2
  AModifier := '';
  Coeff := '';
  Indice := '';
  LibEmp := '';
  Niveau := '';
  Qualif := '';
  StDif := '';
//FIN PT2

  if (WW <> nil) then
  begin
{PT2
   if Q_Mul <> NIL then
      Q_Mul.Liste  := 'PGMODIFCOEFFSAL';
}
    StEvol := 'SELECT *' +
      ' FROM EVOLUTIONSAL';
    QRechEvol := OpenSql(StEvol, TRUE);
    if (not (QRechEvol.Eof)) then
    begin
      TEvol := TOB.Create('Les paramètres evolution', nil, -1);
      TEvol.LoadDetailDB('EVOLUTIONSAL', '', '', QRechEvol, False);
      Amodifier := TEvol.detail[0].GetValue('PVL_EVOLUANT');
      Coeff := TEvol.detail[0].GetValue('PVL_COEFFICIENTP');
      LibEmp := TEvol.detail[0].GetValue('PVL_LIBEMPLOIP');
      Qualif := TEvol.detail[0].GetValue('PVL_QUALIFP');
      Niveau := TEvol.detail[0].GetValue('PVL_NIVEAUP');
      Indice := TEvol.detail[0].GetValue('PVL_INDICEP');
      Ferme(QRechEvol);
    end;
    if (Amodifier = '001') then
    begin
      StDif := ' AND PMC_COEFFICIENT<>PSA_COEFFICIENT';
      if Q_Mul <> nil then
        TFMul(Ecran).SetDBListe('PGMODIFCOEFFSAL');
    end
    else
      if (Amodifier = '002') then
      begin
        StDif := ' AND PMC_QUALIFICATION<>PSA_QUALIFICATION';
        if Q_Mul <> nil then
          TFMul(Ecran).SetDBListe('PGMODIFQUALIF');
      end
      else
        if (Amodifier = '003') then
        begin
          StDif := ' AND PMC_NIVEAU<>PSA_NIVEAU';
          if Q_Mul <> nil then
            TFMul(Ecran).SetDBListe('PGMODIFNIVEAU');
        end
        else
          if (Amodifier = '004') then
          begin
            StDif := ' AND PMC_INDICE<>PSA_INDICE';
            if Q_Mul <> nil then
              TFMul(Ecran).SetDBListe('PGMODIFINDICE');
          end;

    if (Coeff = 'X') then
      StDif := StDif + ' AND PMC_COEFFICIENT=PSA_COEFFICIENT';

    if (LibEmp = 'X') then
      StDif := StDif + ' AND PMC_LIBELLEEMPLOI=PSA_LIBELLEEMPLOI';

    if (Qualif = 'X') then
      StDif := StDif + ' AND PMC_QUALIFICATION=PSA_QUALIFICATION';

    if (Niveau = 'X') then
      StDif := StDif + ' AND PMC_NIVEAU=PSA_NIVEAU';

    if (Indice = 'X') then
      StDif := StDif + ' AND PMC_INDICE=PSA_INDICE';
//FIN PT2

    WW.Text := ' (PSA_DATESORTIE is null OR' +
      ' PSA_DATESORTIE <= "' + UsDateTime(IDate1900) + '"' + StDif + ')';
  end;
end;

initialization
  registerclasses([TOF_PGMULMODIFCOEFFSAL]);
end.

