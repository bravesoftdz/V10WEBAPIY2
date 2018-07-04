{
PT1   30/11/2004 MF V_60 Correction traitement maintien qd champ catégorie renseigné.
                           + insertion commentaire pour maintien
PT2   29/12/2004 MF V_60 correction erreur SQL DB2
PT3   04/01/2004 MF V_60 Modif : la rubrique à utiliser pour le maintien est
                         récupérée de la Tob_Maintien
PT4   29/04/2005 MF V_602 FQ 12234 Nvelle focntion pour annulation des commentaires de maintien
PT5   01/06/2005 MF V_602 FQ 12342 Pas de commentaire qd nbre de jours maintenus = 0
                          (cas carence > nb jours calendaires d'absence)
PT6   23/01/2006 SB V_65 FQ 10866 Ajout clause predefini motif d'absence
PT7   20/02/2006 MF V_65 Traitement des lignes de garantie en testant le champ
                         PMT_TYPEMAINTIEN au lieu du champ PMT_CARENCE = 9999
PT8   23/02/2006 MF V_65 Traitement de la garantie
PT9   28/02/2006 MF V-65 Traitement de la garantie (suite)
PT10  25/04/2006 MF V_65 Commentaire de la garantie présent uniquement si n° rub
                         de maintien = n ° rub de garantie
PT11  05/06/2008 MF V_82  FQ 459 et FQ 15466 : mise au point insertion des commentaires                        
}
unit PGMaintien;

interface
uses
  HCtrls,
{$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
{$ENDIF}
//unused  EntPaie,
  Hent1,
//unused  HMsgBox,
  P5Util,
//unused  PgCongesPayes,
  SysUtils,
  UTOB;

function RecupereMaintien(Datef: tdatetime; Salarie, Action: string; var TOB_Maintien : TOB): boolean;
procedure IntegreMaintienDansPaye(Tob_Rub, Salarie: tob; const DateD, DateF: tdatetime; const Action: string; const subrogation: boolean;const ChampCateg : string;var TOB_Maintien : TOB);
procedure MajDateIntegr(var t: tob; const Dated, Datef: tdatetime);
procedure EnleveCommMaintien(Salarie, Tob_rub: tob; const Arub, Natrub : string; const dated, datef: tdatetime); // PT4

var
  Anciennete: WORD;
  Convention: string;
  StWhere: string;
  Categorie: string;
implementation

function RecupereMaintien(Datef: tdatetime; Salarie, Action: string;var TOB_Maintien : TOB): boolean;

var
  st: string;
  t: Tob;
  Q: TQuery;
  MaDate: TDateTime;
begin
  result := False;
  if Action = 'M' then
  begin
    st := 'SELECT * FROM MAINTIEN WHERE' +
      ' PMT_SALARIE = "' + Salarie + '"' +
      ' AND PMT_DATEFINABS <= "' + usdatetime(Datef) + '"' +
      ' AND (PMT_DATEFIN= "' + usdatetime(2) + '"' +
      ' OR PMT_DATEFIN="' + usdatetime(Datef) + '")';

    StWhere := ' PMT_SALARIE = "' + Salarie + '"' +
      ' AND PMT_DATEFINABS <= "' + usdatetime(Datef) + '"' +
      ' AND (PMT_DATEFIN= "' + usdatetime(2) + '"' +
      ' OR PMT_DATEFIN="' + usdatetime(Datef) + '")';
  end
  else
  begin
    st := 'SELECT * FROM MAINTIEN WHERE' +
      ' PMT_SALARIE = "' + Salarie + '"' +
      ' AND PMT_DATEFINABS <= "' + usdatetime(Datef) + '"' +
      ' AND PMT_DATEFIN <= "' + usdatetime(2) + '"';

    StWhere := ' PMT_SALARIE = "' + Salarie + '"' +
      ' AND PMT_DATEFINABS <= "' + usdatetime(Datef) + '"' +
      ' AND (PMT_DATEFIN= "' + usdatetime(2) + '"' +
      ' OR PMT_DATEFIN="' + usdatetime(Datef) + '")';
  end;

  Q := OpenSql(st, TRUE);
  if not Q.eof then
  begin
    Tob_Maintien := Tob.create('MAINTIEN à intégrer', nil, -1);
    Tob_Maintien.LoadDetailDB('MAINTIEN', '', '', Q, False);
    MaDate := Idate1900;
    { Si maintien non intégré existe renvoie true pour réintégration paye }
    if Tob_Maintien.Findfirst(['PMT_DATEFIN'], [Madate], false) <> nil then
      result := true;
  end
  else
    if Tob_Maintien <> nil then FreeAndNil(Tob_Maintien);
  Ferme(Q);


  if (Action = 'M') and (Tob_Maintien <> nil) and (result) then
    { On réintègre tous les maintiens même ceux déjà intégré non!!}
  begin
//    t := Tob_Maintien.Findfirst([''], [''], false);
    t := Tob_Maintien.Findfirst(['PMT_DATEFIN'], [Datef], false);
    while t <> nil do
    begin
      T.free;
      T := Tob_Maintien.Findnext(['PMT_DATEFIN'], [Datef], false)
    end;
  end;
end;


procedure IntegreMaintienDansPaye(Tob_Rub, Salarie: tob; const DateD, DateF: tdatetime; const Action: string; const subrogation: boolean; Const ChampCateg : string;var TOB_Maintien : TOB);
var
  Q                                        : tQuery;
  st                                       : string;
  t: tob;
  TypeAbs, TypeAbsP, Aliment, Natrub       : string;
  RubMaintien, RubMaintienP                : string;
  TypemaintienP                            : string;  //PT10
  ValMaintien                              : double;
  i                                        : integer; //PT1
begin
  if BullCompl = 'X' then exit;

  ValMaintien := 0;
  i := 1; //PT1
  if TOB_Maintien = nil then exit;
// d  PT9
  TOB_Maintien.Detail.Sort('PMT_TYPEMAINTIEN;PMT_DATEDEBUTABS');
// f PT9
  t := TOB_Maintien.findfirst([''], [''], false);
  if t = nil then exit;

// d PT3
  RubMaintien := t.GetValue('PMT_RUBRIQUE');
// f PT3
  RubMaintienP := RubMaintien;
  TypemaintienP := t.GetValue('PMT_TYPEMAINTIEN');    //PT10

  Natrub := 'AAA';

  if (Action = 'M') then
  begin
    EnleveCommMaintien(Salarie, Tob_rub, RubMaintien, Natrub, DateD, Datef); // PT4
  end;

  while t <> nil do
  begin
    st := 'SELECT PMA_TYPEABS FROM MOTIFABSENCE WHERE ##PMA_PREDEFINI## PMA_MOTIFABSENCE = "' +   { PT6 }
      t.getvalue('PMT_TYPECONGE') + '"';
    Q := opensql(st, True);
    if not Q.eof then
      TypeAbs := Q.findfield('PMA_TYPEABS').AsString
// d PT8
    else
      TypeAbs := t.getvalue('PMT_TYPECONGE');
// f PT8
    ferme(Q);
    if (TypeAbs <> TypeAbsP) and (RubMaintien <> RubMaintienP) then  //PT1
    begin
// d PT8
//      ValMaintien := ValMaintien + t.GetValue('PMT_MTMAINTIEN');
//      IntegreRub(tob_rub, Salarie, typeAbsP, Aliment, DateD, DateF, ValMaintien, Natrub, RubMaintien, False, True);
// f PT8
      IntegreRub(tob_rub, Salarie, typeAbsP, Aliment, DateD, DateF, ValMaintien, Natrub, RubMaintienP, False, True);
      typeAbsP := TypeAbs;

      ValMaintien := 0;

//d PT3
      RubMaintien := t.GetValue('PMT_RUBRIQUE');
//f PT3
      if RubMaintien <> RubMaintienP then
      begin
        EnleveCommMaintien(Salarie, Tob_rub, RubMaintien, Natrub, DateD, Datef);  // PT4
// d PT10
        // le commentaire de la garantie n'existe que si n° rub maintien = n° rub garantie
{PT11
        if (TypeMaintienP <> t.getValue('PMT_TYPEMAINTIEN')) then}
        if (TypeMaintienP = 'GAR') and (TypeMaintienP <> t.getValue('PMT_TYPEMAINTIEN')) then
          EnleveCommMaintien(Salarie, Tob_rub, RubMaintienP, Natrub, DateD, Datef);
// f PT10
        RubMaintienP := RubMaintien;    
        TypemaintienP := t.GetValue('PMT_TYPEMAINTIEN');    // PT11
      end;
    end;
//d PT1
// d  PT5 FQ12342
// d PT7
//  if (t.GetValue('PMT_NBJMAINTIEN') <> 0) or (t.GetValue('PMT_CARENCE') = 9999) then
    if (t.GetValue('PMT_NBJMAINTIEN') <> 0) or (t.GetValue('PMT_TYPEMAINTIEN') = 'GAR') then
// f PT7
// f  PT5 FQ12342
      EcritCommMaint(Tob_rub, Salarie, t, i, DateD, Datef, RubMaintien, NatRub);

//f PT1
    ValMaintien := ValMaintien + t.GetValue('PMT_MTMAINTIEN');
    MajDateIntegr(t, DateD, DateF);
    t := Tob_Maintien.findnext([''], [''], false);
//d PT8
    if (t <> nil) then
        RubMaintien := t.GetValue('PMT_RUBRIQUE');
// f PT8
  end;
  // Ecriture du dernier
// d PT3
  IntegreRub(tob_rub, Salarie, typeAbsP, 'MON', DateD, DateF, ValMaintien, Natrub, RubMaintien, False, True);
// f PT3
end;
// d PT4
procedure EnleveCommMaintien(Salarie, Tob_rub: tob; const Arub, Natrub : string; const dated, datef: tdatetime);
var
  t: tob;
begin
  // PT14 : 23/10/01 : V562 : PH Rajout traitement bulletion complémentaire et dates édition
  if BullCompl = 'X' then exit;

  // mv on vire les libelles de la tob_rub créés précédemment
  T := Tob_Rub.FindFirst(['PHB_NATURERUB', 'PHB_DATEDEBUT', 'PHB_DATEFIN', 'PHB_SALARIE'],
    [Natrub, DateD, DateF, Salarie.GetValeur(iPSA_SALARIE)], TRUE);
  while T <> nil do
  begin
    if ((copy(T.GetValue('PHB_RUBRIQUE'), 1, length(ARub) + 1) = ARub + '.') and
      (T.GetValue('PHB_ORIGINELIGNE') = 'MAI')) then
      T.free;
    T := Tob_Rub.FindNext(['PHB_NATURERUB', 'PHB_DATEDEBUT', 'PHB_DATEFIN', 'PHB_SALARIE'],
      [Natrub, DateD, DateF, Salarie.GetValeur(iPSA_SALARIE)], TRUE);
  end;
end;
// f PT4

procedure MajDateIntegr(var t: tob; const DateD, datef: tdatetime);
begin
  if T = nil then exit;
  T.putvalue('PMT_DATEDEBUT', DateD);
  T.putvalue('PMT_DATEFIN', datef);
end;
end.

