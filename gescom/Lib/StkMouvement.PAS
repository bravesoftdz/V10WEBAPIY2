{***********UNITE*************************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 23/04/2003
Modifié le ... :   /  /
Description .. : Fonctions & procédure métier
Mots clefs ... : STKMOUVEMENT

Sql de recalcul des stock:
UPDATE DISPO
SET GQ_RESERVEFOU = IsNull((SELECT SUM(ABS(GSM_QPREVUE)-ABS(GSM_PHYSIQUE)) FROM STKMOUVEMENT WHERE GSM_STKTYPEMVT="ATT" AND GSM_ETATMVT<>"SOL" AND GSM_ARTICLE=GQ_ARTICLE AND GSM_DEPOT=GQ_DEPOT), 0),
    GQ_RESERVECLI = IsNull((SELECT SUM(ABS(GSM_QPREVUE)-ABS(GSM_PHYSIQUE)) FROM STKMOUVEMENT WHERE GSM_STKTYPEMVT="RES" AND GSM_ETATMVT<>"SOL" AND GSM_ARTICLE=GQ_ARTICLE AND GSM_DEPOT=GQ_DEPOT), 0),
    GQ_PHYSIQUE = IsNull((SELECT SUM(GSM_PHYSIQUE) FROM STKMOUVEMENT WHERE GSM_STKTYPEMVT="PHY" AND GSM_ARTICLE=GQ_ARTICLE AND GSM_DEPOT=GQ_DEPOT), 0)

Delete JLS PF
delete StkMouvement where GSM_ARTICLE = "JLS PF                           X";
delete Dispo        where GQ_ARTICLE  = "JLS PF                           X";
delete DispoDetail  where GQD_ARTICLE = "JLS PF                           X";
delete StkMouvement where GSM_ARTICLE = "JLS SM                           X";
delete Dispo        where GQ_ARTICLE  = "JLS SM                           X";
delete DispoDetail  where GQD_ARTICLE = "JLS SM                           X";
delete StkMouvement where GSM_ARTICLE = "JLS A                            X";
delete Dispo        where GQ_ARTICLE  = "JLS A                            X";
delete DispoDetail  where GQD_ARTICLE = "JLS A                            X";
delete StkMouvement where GSM_ARTICLE = "JLS B                            X";
delete Dispo        where GQ_ARTICLE  = "JLS B                            X";
delete DispoDetail  where GQD_ARTICLE = "JLS B                            X";
delete StkMouvement where GSM_ARTICLE = "JLS C                            X";
delete Dispo        where GQ_ARTICLE  = "JLS C                            X";
delete DispoDetail  where GQD_ARTICLE = "JLS C                            X";


Select
Select * from StkMouvement where GSM_ARTICLE = "JLS PF                           X"
*****************************************************************}
unit StkMouvement;

interface

uses
  Windows,
  Classes,
  SysUtils,
  Hctrls,
  Hent1,
  uTob,
  uTom,
  Dialogs,
  Menus,
  Graphics,
  SaisieList,
  uTableFiltre,
  messages,
  Hmsgbox,

{$IFNDEF EAGLCLIENT}
  DBtables,
  Db,
  FE_Main,
{$ELSE}
  MainEAGL,
{$ENDIF}

  wCommuns
  ;

{$IFDEF STK}
Type
  tCleGSM = record
    StkTypeMVT  : string;
    QualifMVT   : String;
    Identifiant : Integer;
  end;
  TActionGSM = Class
    TobData : Tob;
  public
    procedure DoCreate;
    procedure DoModify;
    procedure DoDelete;
    procedure DoAddToRealise;
  end;

function ModifyGSM(TobData: Tob): boolean;

{ Call }
procedure CallPhyGSM(Argument, Flux: string; Titre: string = '');
procedure CallAttGSM(Argument, Flux: string);
procedure CallResGSM(Argument, Flux: string);
procedure CallProGSM(Argument, Flux: string);
procedure CallFicRefOrigine(Reforigine: String);

procedure CallFicGSM(CleGSM: TCleGSM; Flux: string; Action: String; UniteSto: String =''; UniteVte: String =''; UniteAch: String =''; UnitePro: String =''; UniteCon: String ='');

{$ENDIF}

const
	TableName = 'STKMOUVEMENT';

implementation

Uses
  EntGP,
  StkNature,
  UtilArticle,
  wOrdreLig,
  wOrdreBes,
  DispoDetail,
  Dispo
  ;

{$IFDEF STK}

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 24/04/2003
Modifié le ... : 24/04/2003
Description .. : Modification de GSM
Suite ........ : Ne s'applique qu'aux mouvements ATT et RES
Suite ........ : A chaque TobData.Detail peut correspondre x TobGSM
Suite ........ : (affectations)
Suite ........ :
Suite ........ : JLS 24/04/2002: Attention je n'ai pas géré les affectations.
Suite ........ : Aujourd'hui, je n'ai qu'une ligne dans TobGSM
Mots clefs ... :
*****************************************************************}
function ModifyGSM(TobData: Tob): boolean;
var
  WithProgressForm: Boolean;
  iData           : integer;
  sql             : string;
  TobGSM          : Tob;

  function ModifyTobGSM: boolean;
  var
    iGSM, iChamp: integer;
    TomGSM      : Tom;
  begin
    TomGSM := CreateTOM(TableName, nil, false, true);
    try
      for iGSM := 0 to TobGSM.Detail.Count - 1 do
      begin
        { Sauvegarde de la Tob }
        TomGSM.LoadBufferAvantModif(TobGSM.Detail[iGSM]);

        { Data }
        for iChamp := 1000 to (1000 + TobData.Detail[iData].ChampsSup.Count - 1) do
        begin
          tobGSM.Detail[iGSM].P(TobData.Detail[iData].GetNomChamp(iChamp), TobData.Detail[iData].G(TobData.Detail[iData].GetNomChamp(iChamp)));
        end;

        { UpDateRecord }
        tobGSM.Detail[iGSM].AddChampSupValeur('IKC', 'M', false);
        TomGSM.VerifTOB(tobGSM.Detail[iGSM]);
      end;
      Result := TobGSM.UpDateDb;
    finally
      TomGSM.free;
    end;
  end;

begin
  Result := false;
  if Assigned(TobData) then
  begin
    WithProgressForm := TobData.Detail.Count > 1;
    if WithProgressForm then wInitProgressForm(nil, TraduireMemoire('Modification mouvement'), '', TobData.Detail.Count, False, True);
    try
      for iData := 0 to TobData.Detail.Count - 1 do
      begin
        if WithProgressForm then wMoveProgressForm;

        { On charge la TobGSM avec la ou les lignes correspondant à TobData }
        Sql := 'SELECT *'
             + ' FROM ' + TableName
             + ' WHERE GSM_REFORIGINE="' + TobData.Detail[iData].G('GSM_REFORIGINE') + '"'
             + ' ORDER BY GSM_PRIORITEAFF'
             ;
        TobGSM := Tob.Create('GSM', nil, -1);
        try
          Result := wLoadTobFromSql(TableName, Sql, TobGSM) and ModifyTobGSM;
        finally
          TobGSM.free;
        end;
      end;
    finally
      if WithProgressForm then wFiniProgressForm;
    end;
  end;
end;

{ TActionGSM }
{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 23/04/2003
Modifié le ... : 24/06/2003
Description .. : Création de mouvement de stock
Suite ........ : TobData: Tob virtuelle contenant autant de détail que de
Suite ........ : mouvements à générer.
Suite ........ : Chaque mouvement à générer peut donner x mouvements
Suite ........ : réels.
Mots clefs ... :
*****************************************************************}
procedure TActionGSM.DoCreate;
var
  withProgressForm: boolean;
  iGSM, iChamp    : integer;
  TomGSM          : Tom;
  TobGSM          : Tob;
begin
  if Assigned(TobData) then
  begin
    TomGSM := CreateTOM(TableName, nil, false, true);
    TobGSM := Tob.Create(TableName, nil, -1);
    WithProgressForm := TobData.Detail.Count > 1;
    try
      if WithProgressForm then wInitProgressForm(nil, TraduireMemoire('Création mouvement'), '', TobData.Detail.Count, False, True);
      for iGSM := 0 to TobData.Detail.Count - 1 do
      begin
        if WithProgressForm then wMoveProgressForm;

        { NewRecord }
        tomGSM.InitTOB(tobGSM);

        { TobData }
        TobData.Detail[iGSM].DelChampSup('GSM_IDENTIFIANT', false);
        for iChamp := 1000 to (1000 + TobData.Detail[iGSM].ChampsSup.Count - 1) do
        begin
          tobGSM.P(TobData.Detail[iGSM].GetNomChamp(iChamp), TobData.Detail[iGSM].G(TobData.Detail[iGSM].GetNomChamp(iChamp)));
        end;

        { UpDateRecord }
        tobGSM.AddChampSupValeur('IKC', 'C', false);
        if tomGSM.VerifTOB(tobGSM) then
        begin
          tobGSM.InsertDB(nil);

          {AfterUpDateRecord}
          tomGSM.AfterVerifTOB(tobGSM);
        end
        else
        begin
          TobData.Detail[iGSM].AddChampSupValeur('Error', tobGSM.GetValue('Error'), false);
        end;
      end;
    finally
      if WithProgressForm then wFiniProgressForm;
      TomGSM.free;
      TobGSM.free;
    end;
  end;
end;

procedure TActionGSM.DoDelete;
  function GetWhere: string;
  var
    iData: integer;
  begin
    for iData := 0 to TobData.Detail.Count - 1 do
    begin
      if iData > 0 then
        Result := Result + ' AND ';

      Result := Result + 'GSM_REFORIGINE="' + TobData.Detail[iData].G('GSM_REFORIGINE') + '"';
    end;
  end;
begin
  if Assigned(TobData) then
    wDeleteTable(TableName, GetWhere);
end;

procedure TActionGSM.DoModify;
var
  WithProgressForm: Boolean;
  iData           : integer;
  sql             : string;
  TobGSM          : Tob;

  procedure ModifyTobGSM;
  var
    iGSM, iChamp: integer;
    TomGSM      : Tom;
  begin
    TomGSM := CreateTOM(TableName, nil, false, true);
    try
      for iGSM := 0 to TobGSM.Detail.Count - 1 do
      begin
        { Sauvegarde de la Tob }
        TomGSM.LoadBufferAvantModif(TobGSM.Detail[iGSM]);

        { Data }
        for iChamp := 1000 to (1000 + TobData.Detail[iData].ChampsSup.Count - 1) do
        begin
          tobGSM.Detail[iGSM].P(TobData.Detail[iData].GetNomChamp(iChamp), TobData.Detail[iData].G(TobData.Detail[iData].GetNomChamp(iChamp)));
        end;

        { UpDateRecord }
        tobGSM.Detail[iGSM].AddChampSupValeur('IKC', 'M', false);
        if TomGSM.VerifTOB(tobGSM.Detail[iGSM]) then
        begin
          tobGSM.Detail[iGSM].UpDateDb;

          { OnAfterUpDateRecord }
          TomGSM.AfterVerifTOB(tobGSM.Detail[iGSM])
        end;
      end;
    finally
      TomGSM.free;
    end;
  end;

begin
  if Assigned(TobData) then
  begin
    WithProgressForm := TobData.Detail.Count > 1;
    if WithProgressForm then wInitProgressForm(nil, TraduireMemoire('Modification mouvement'), '', TobData.Detail.Count, False, True);
    try
      for iData := 0 to TobData.Detail.Count - 1 do
      begin
        if WithProgressForm then wMoveProgressForm;

        if TobData.Detail[iData].G('GSM_QUALIFMVT') <> '' then
        begin
          { On charge la TobGSM avec la ou les lignes correspondant à TobData }
          Sql := 'SELECT *'
               + ' FROM ' + TableName
               + ' WHERE GSM_STKTYPEMVT="' + GetFieldFromGSN('GSN_STKTYPEMVT', TobData.Detail[iData].G('GSM_QUALIFMVT')) + '"'
               + ' AND GSM_REFORIGINE="' + TobData.Detail[iData].G('GSM_REFORIGINE') + '"'
               + ' ORDER BY GSM_PRIORITEAFF'
               ;
          TobGSM := Tob.Create('GSM', nil, -1);
          try
            if wLoadTobFromSql(TableName, Sql, TobGSM) then
              ModifyTobGSM;
          finally
            TobGSM.free;
          end;
        end;
      end;
    finally
      if WithProgressForm then wFiniProgressForm;
    end;
  end;
end;

procedure TActionGSM.DoAddToRealise;
var
  WithProgressForm: Boolean;
  iData           : integer;
  sql             : string;
  TobGSM          : Tob;

  procedure AddToRealiseTobGSM;
  var
    iGSM  : integer;
    TomGSM: Tom;
  begin
    TomGSM := CreateTOM(TableName, nil, false, true);
    try
      for iGSM := 0 to TobGSM.Detail.Count - 1 do
      begin
        { Sauvegarde de la Tob }
        TomGSM.LoadBufferAvantModif(TobGSM.Detail[iGSM]);

        { Cumul }
        tobGSM.Detail[iGSM].P('GSM_PHYSIQUE', tobGSM.Detail[iGSM].G('GSM_PHYSIQUE') + TobData.detail[iData].G('GSM_PHYSIQUE'));

        { Remettre l'état correctement }
        if (tobGSM.Detail[iGSM].G('GSM_ETATMVT')= 'SOL') and (Abs(TomGSM.GetFieldAvantModif('GSM_PHYSIQUE')) > Abs(tobGSM.Detail[iGSM].G('GSM_PHYSIQUE'))) then
          tobGSM.Detail[iGSM].P('GSM_ETATMVT', 'FER');


        { UpDateRecord }
        tobGSM.Detail[iGSM].AddChampSupValeur('IKC', 'M', false);
        if TomGSM.VerifTOB(tobGSM.Detail[iGSM]) then
        begin
          tobGSM.Detail[iGSM].UpDateDb;

          { OnAfterUpDateRecord }
          TomGSM.AfterVerifTOB(tobGSM.Detail[iGSM])
        end;
      end;

    finally
      TomGSM.free;
    end;
  end;

  function GetQualifMvtRef: string;
  begin
    Result := StkNature.GetQualifMvtRef(TobData.Detail[iData].G('GSM_QUALIFMVT'));
  end;
begin
  if Assigned(TobData) then
  begin
    WithProgressForm := TobData.Detail.Count > 1;
    if WithProgressForm then wInitProgressForm(nil, TraduireMemoire('Cumul du réalisé'), '', TobData.Detail.Count, False, True);
    try
      for iData := 0 to TobData.Detail.Count - 1 do
      begin
        if WithProgressForm then wMoveProgressForm;

        { On charge la TobGSM avec la ou les lignes correspondant à TobData }
        Sql := 'SELECT *'
             + ' FROM ' + TableName
             + ' WHERE GSM_STKTYPEMVT="' + GetFieldFromGSN('GSN_STKTYPEMVT', GetQualifMvtRef) + '"'
             + ' AND GSM_QUALIFMVT="' + GetQualifMvtRef + '"'
             + ' AND GSM_REFORIGINE="' + TobData.Detail[iData].G('GSM_REFORIGINE') + '"'
             + ' ORDER BY GSM_PRIORITEAFF'
             ;
        TobGSM := Tob.Create('GSM', nil, -1);
        try
          if wLoadTobFromSql(TableName, Sql, TobGSM) then
            AddToRealiseTobGSM;
        finally
          TobGSM.free;
        end;
      end;
    finally
      if WithProgressForm then wFiniProgressForm;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 30/07/2003
Modifié le ... :   /  /
Description .. : Appel de la fiche des mouvemenst de stock
Mots clefs ... :
*****************************************************************}
procedure CallPhyGSM(Argument, Flux: string; Titre: string = '');
  function GetRange: string;
  begin
    Result := 'GSM_QUALIFMVT=' + GetArgumentValue(Argument, 'QUALIFMVT')
            + ';GSM_CODEARTICLE=' + GetArgumentValue(Argument, 'CODEARTICLE')
            + ';GSM_DEPOT=' + GetArgumentValue(Argument, 'DEPOT')
            + ';GSM_REFORIGINE=' + GetArgumentValue(Argument, 'REFORIGINE')
            + ';GSM_SERIEINTERNE=' + GetArgumentValue(Argument, 'SERIEINTERNE')

  end;
  function GetArgument: string;
  begin
    Result := 'FLUX=' + Flux
            + ';TITRE=' + Titre
  end;
begin
  wForceCumulsOnViewFields('GSMPHYSIQUE', 'PHYSIQUESTO;PHYSIQUEVTE;PHYSIQUEACH;PHYSIQUEPRO;PHYSIQUECON;');
  AglLanceFiche('GC', 'GSMPHYSIQUE_MUL', GetRange, '', 'ACTION=CONSULTATION;' + GetArgument);
end;

procedure CallAttGSM(Argument, Flux: string);
  function GetRange: string;
  begin
    Result := 'GSM_QUALIFMVT=' + GetArgumentValue(Argument, 'QUALIFMVT')
            + ';GSM_CODEARTICLE=' + GetArgumentValue(Argument, 'CODEARTICLE')
            + ';GSM_DEPOT=' + GetArgumentValue(Argument, 'DEPOT')
            + ';GSM_REFORIGINE=' + GetArgumentValue(Argument, 'REFORIGINE')
  end;
  function GetArgument: string;
  begin
    Result := 'FLUX=' + Flux
  end;
begin
  wForceCumulsOnViewFields('GSMATTENDUS', 'QRESTESTO;QRESTEVTE;QRESTEACH;QRESTEPRO;QRESTECON;');
  AglLanceFiche('GC', 'GSMATTENDUS', GetRange, '', 'ACTION=MODIFICATION;' + GetArgument);
end;

procedure CallResGSM(Argument, Flux: string);
  function GetRange: string;
  begin
    Result := 'GSM_QUALIFMVT=' + GetArgumentValue(Argument, 'QUALIFMVT')
            + ';GSM_CODEARTICLE=' + GetArgumentValue(Argument, 'CODEARTICLE')
            + ';GSM_DEPOT=' + GetArgumentValue(Argument, 'DEPOT')
            + ';GSM_REFORIGINE=' + GetArgumentValue(Argument, 'REFORIGINE')
  end;
  function GetArgument: string;
  begin
    Result := 'FLUX=' + Flux
  end;
begin
  wForceCumulsOnViewFields('GSMRESERVES', 'QRESTESTO;QRESTEVTE;QRESTEACH;QRESTEPRO;QRESTECON;');
  AglLanceFiche('GC', 'GSMRESERVES', GetRange, '', 'ACTION=MODIFICATION;' + GetArgument);
end;

procedure CallProGSM(Argument, Flux: string);
  function GetRange: string;
  begin
    Result := ';GSM_CODEARTICLE=' + GetArgumentValue(Argument, 'CODEARTICLE')
            + ';GSM_DEPOT=' + GetArgumentValue(Argument, 'DEPOT')
   end;
  function GetArgument: string;
  begin
    Result := 'FLUX=' + Flux
  end;
begin
  AglLanceFiche('GC', 'GSMPROJETES', GetRange, '', 'ACTION=MODIFICATION;' + GetArgument);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 22/05/2003
Modifié le ... :   /  /
Description .. : Appel de la fiche correspondant à la RefOrigine
Mots clefs ... :
*****************************************************************}
procedure CallFicRefOrigine(Reforigine: String);
  function GetPrefixe: string;
  begin
    Result := Copy(RefOrigine, 1, Pos('~', RefOrigine)-1);
  end;
begin
  if RefOrigine <> '' then
  begin
    if      GetPrefixe = 'WOL' then wCallFicWOL(GetCleWOL(RefOrigine))
    else if GetPrefixe = 'WOB' then wCallFicWOB(GetCleWOB(RefOrigine))
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Dominique Sclavopoulos
Créé le ...... : 17/06/2003
Modifié le ... :   /  /
Description .. : Appel de la fiche des mouvements
Mots clefs ... :
*****************************************************************}
procedure CallFicGSM(CleGSM: TCleGSM; Flux: string; Action: String; UniteSto, UniteVte, UniteAch, UnitePro, UniteCon: String);
begin
  AGLLanceFiche('GC', 'GSMPHYSIQUE_FIC', '', CleGSM.STKTypeMVT + ';' + CLeGSM.QualifMvt + ';' + IntToStr(CleGSM.Identifiant), 'MONOFICHE;ACTION=' + Action + ';FLUX=' + Flux + ';UNITESTO=' + UniteSto + ';UNITEVTE=' + UniteVte + ';UNITEACH=' + UniteAch + ';UNITEPRO=' + UnitePro + ';UNITECON=' + UniteCon); 
end;

{$ENDIF}
end.
