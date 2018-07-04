{***********UNITE*************************************************
Auteur  ...... : BRL
Créé le ...... : 27/04/2004                                                                 
Modifié le ... : 27/04/2004
Description .. : Source TOF de la fiche : BTPAIEGENER
Mots clefs ... : TOF;BTPAIEGENER
*****************************************************************}
unit UTofBTPaieGener;

interface

uses {$IFDEF VER150} variants,{$ENDIF} StdCtrls, Controls, Classes,
  {$IFDEF EAGLCLIENT}
  eMul, Maineagl,
  {$ELSE}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} mul, FE_Main,
  {$ENDIF}
  forms, sysutils, GCMZSUtil, Utob, UtilMultrt, Windows, Hstatus, HTB97,
  ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF,
  Dicobtp, SaisUtil, EntGC, M3FP, ParamSoc,
  UtilRessource, UTofBTPaieGenerChoix,
  utilpgi, AglInit, UtilGc, TraducAffaire, UtofAfTraducChampLibre,
  VentilCpta,Ent1,CBPPath;

type
  TOF_BTPAIEGENER = class(TOF_AFTRADUCCHAMPLIBRE)
    procedure OnArgument(S: string); override;
    procedure OnClose; override;

  private
    FlogName,CheminDest: string;
    FLogList: TStringList;
    procedure BEffaceFicLogOnClick(Sender: TObject);
    procedure BFicLogOnClick(Sender: TObject);
    procedure EcrireFichierLog;
    procedure Log(Msg: string);

    procedure BalanceLaGen;
    procedure Gen;
    procedure TOBCopieChamp(FromTOB, ToTOB: TOB);
    procedure GenereSection(Etab: string; TobActivite, TOBArticles, TOBTiers, TOBAffaires: TOB);
  end;

Procedure BTLanceFiche_ActivGenerPaie;

const
  TexteMessage: array[1..19] of string = (
    {1}'Confirmez-vous la suppression du fichier '
    {2}, 'Le fichier n''existe pas : '
    {3}, 'Vous devez positionner les paramètres société de ressource d''envoi paie '
    {4}, ' - Pas d''activité pour cette ressource '
    {5}, '- Génération des lignes d''activité vers la Paie le '
    {6}, 'Période Paie du '
    {7}, 'Période Activité du '
    {8}, 'Fichier paie n°'
    {9}, 'Exercice inconnu : '
    {10}, ' - Pas de code salarié associé pour cette ressource  '
    {11}, ' - Le code salarié %s associé à cette ressource n''existe pas '
    {12}, 'Activité de la ressource : %s - Code salarié : '
    {13}, 'les lignes générées dans la paie n''ont pas été marquées dans l''activité'
    {14}, 'Fin de la préparation lignes activitées pour %s salarié(s) sur '
    {15}, ' anomalie(s) - consultez le fichier '
    {16}, 'Pas d''activité à traiter pour les ressources sur la période du '
    {17}, 'Impossible d''accèder au fichier de log .'
    {18}, 'Vous n''avez pas sélectionné les ressources à générer dans la paie.'
    {19}, 'Vous devez faire le paramètrage lien paie pour associer les rubriques aux prestations.'
    );

implementation
uses ShellAPI;

const Sep = '|';
type

  TAna = class
    Axe: string;
    Section: string;
    Libelle: string;
    Value: Double;
    Pourcent: Double;
    NumVen: integer;
  end;

  TAnaList = class
  private
    FAnaList: array of TAna;
    function GetSection(Ax, Sect, LibSect: string): TAna;
    function GetAna(index: integer): TAna;
    function GetCount: integer;

  public
    constructor create;
    destructor destroy; override;

    procedure Pourcentize(LAxe : string);

    property Section[Ax, Sect, LibSect: string]: TAna read GetSection;
    property Ana[index: integer]: TAna read GetAna;
    property Count: integer read GetCount;

  end;

  TRub = class
    CodeRub: string;
    Libelle: string;
    Qte: Double;
    Montant: Double;
    GotQte: boolean;
    GotMontant: boolean;
    Ventiled: boolean;
    VentAnal: TAnaList; // Note pour plus tard : trouver un autre nom (celui-là pue un peu)

  public
    procedure PourcentRubrique(LAxe : String);
    constructor create;
    destructor destroy; override;

  end;

  TRubList = class
  private
    FRubList: array of TRub;
    function GetRubByCode(index: string): TRub;
    function GetRub(index: integer): TRub;
    function GetCount: integer;

  public
    constructor create;
    destructor destroy; override;

    property Rubrik[index: string]: TRub read GetRubByCode;
    property Rub[index: integer]: TRub read GetRub;
    property Count: integer read GetCount;
  end;

constructor TAnaList.create;
begin
  inherited;
  SetLength(FAnaList, 0);
end;

destructor TAnaList.destroy;
var i: integer;
begin
  for i := Low(FAnaList) to High(FAnaList) do FAnaList[i].Free;
  SetLength(FAnaList, 0);
  inherited;
end;

procedure TAnaList.Pourcentize (LAxe : String);
var i: integer;
	Total : double;
  dpourc: double;
  Debut,Fin :integer;
begin
  if (High(FAnaList)= -1) then
    Exit;
  dpourc := 0;
  Total := 0;
  Debut := -1; fin := debut;
  for i := Low(FAnaList) to High(FAnaList) do // Totalizage
  begin
  	if FAnaList[i].Axe = LAxe then
    begin
    	if debut= -1 then Debut := i;
      Fin := i;
    	Total := Total + FAnaList[i].Value;
    end;
  end;
  if debut = -1 then Exit; // Aucune ligne à traiter pour cet axe.
	 for i := Debut to fin do // Pourcentizage
   begin
      if Fanalist[I].Axe = LAxe then
      begin
         FanaList[i].Pourcent := Arrondi((Fanalist[I].Value * 100) / Total, 2);
         dpourc := dpourc + FanaList[i].Pourcent;
      end;
   end;

  if debut < fin then
    FAnaList[fin].Pourcent := FAnaList[fin].Pourcent+ (100 - dpourc)
  else
    FAnaList[fin].Pourcent := 100;
end;

procedure TRub.PourcentRubrique(LAxe : String);
var i: integer;
	Total : double;
  dpourc: double;
  Debut,Fin :integer;
begin
  if (High(VentAnal.FAnaList)= -1) then
    Exit;
  dpourc := 0;
  Total := 0;
  Debut := -1; fin := debut;
  for i := Low(VentAnal.FAnaList) to High(VentAnal.FAnaList) do // Totalizage
  begin
  	if VentAnal.FAnaList[i].Axe = LAxe then
    begin
    	if debut= -1 then Debut := i;
      Fin := i;
    	Total := Total + VentAnal.FAnaList[i].Value;
    end;
  end;
  if debut = -1 then Exit; // Aucune ligne à traiter pour cet axe.
	for i := Debut to fin do // Pourcentizage
  begin
     if VentAnal.Fanalist[I].Axe = LAxe then
     begin
        if total <> 0 then VentAnal.FanaList[i].Pourcent := Arrondi((VentAnal.Fanalist[I].Value * 100) / Total, 2);
        dpourc := dpourc + VentAnal.FanaList[i].Pourcent;
     end;
  end;

  if debut < fin then
    VentAnal.FAnaList[fin].Pourcent := VentAnal.FAnaList[fin].Pourcent+ (100 - dpourc)
  else
    VentAnal.FAnaList[fin].Pourcent := 100;
end;

function TAnaList.GetSection(Ax, Sect, LibSect: string): TAna;
var i, nv: integer;
begin
  nv := 1;
  for i := Low(FAnaList) to High(FAnaList) do
    if FAnaList[i].Axe = Ax then
    begin
      inc(nv);
      if FAnaList[i].Section = Sect then
      begin
        result := FAnaList[i];
        exit;
      end;
    end;
  SetLength(FAnaList, Length(FAnaList) + 1);
  FAnaList[High(FAnaList)] := TAna.Create;
  FAnaList[High(FAnaList)].Axe := Ax;
  FAnaList[High(FAnaList)].Section := Sect;
  FAnaList[High(FAnaList)].NumVen := nv;
  FAnaList[High(FAnaList)].Libelle := LibSect;
  result := FAnaList[High(FAnaList)];
end;

function TAnaList.GetAna(index: integer): TAna;
begin
  result := FAnaList[index];
end;

function TAnaList.GetCount: integer;
begin
  result := Length(FAnaList);
end;

constructor TRub.create;
begin
  inherited;
  VentAnal := TAnaList.Create;
end;

destructor TRub.destroy;
begin
  VentAnal.Free;
  inherited;
end;

constructor TRubList.create;
begin
  inherited;
  SetLength(FRubList, 0);
end;

destructor TRubList.destroy;
var i: integer;
begin
  for i := Low(FRubList) to High(FRubList) do FRubList[i].Free;
  SetLength(FRubList, 0);
  inherited;
end;

function TRubList.GetRubByCode(index: string): TRub;
var i: integer;
begin
  for i := Low(FRubList) to High(FRubList) do
    if FRubList[i].CodeRub = index then
    begin
      result := FRubList[i];
      exit;
    end;
  SetLength(FRubList, Length(FRubList) + 1);
  FRubList[High(FRubList)] := TRub.Create;
  FRubList[High(FRubList)].CodeRub := index;
  result := FRubList[High(FRubList)];
end;

function TRubList.GetRub(index: integer): TRub;
begin
  result := FRubList[index];
end;

function TRubList.GetCount: integer;
begin
  result := Length(FRubList);
end;

// -----------------------------------------------------------------------------

procedure TOF_BTPAIEGENER.OnArgument(S: string);
var CC : THValComboBox;
begin
	fMulDeTraitement := true;
  inherited;
  FTableName := 'RESSOURCES';
  CC:=THValComboBox(GetControl('ARS_ETABLISSEMENT')) ; if CC<>Nil then PositionneEtabUser(CC) ;
  TToolBarButton97(GetControl('BFICLOG')).OnClick := BFicLogOnClick;
  TToolBarButton97(GetControl('BEFFACEFICLOG')).OnClick := BEffaceFicLogOnClick;
  SetControlText ('ARS_TYPERESSOURCE', 'SAL');
  SetControlEnabled ('ARS_TYPERESSOURCE', false);
end;

procedure TOF_BTPAIEGENER.OnClose;
begin
  inherited;
end;

procedure TOF_BTPAIEGENER.BEffaceFicLogOnClick(Sender: TObject);
var stDocWord: string;
begin
  stDocWord := GetControlText('FICHIERLOG');
  if (stDocWord = '') then exit;
  if (PGIAskAF(TexteMessage[1] + stDocWord + ' ?', ecran.caption) = mrYes) then
    if not DeleteFile(PChar(stDocWord)) then
      PGIBoxAF(TexteMessage[2] + stDocWord, ecran.caption);
end;

procedure TOF_BTPAIEGENER.BFicLogOnClick(Sender: TObject);
var stDocWord: string;
begin
  stDocWord := GetControlText('FICHIERLOG');
  if (stDocWord = '') then exit;
  if not FileExists(stDocWord)
    then PGIBoxAF(TexteMessage[2] + stDocWord, ecran.caption)
  else ShellExecute(0, 'open', PChar(stDocWord), nil, nil, SW_SHOW);
end;

procedure TOF_BTPAIEGENER.EcrireFichierLog;
var FLogFile: TextFile;
    i :integer;
begin
  FLogName := GetControlText('FICHIERLOG');
  if trim(FLogName) = '' then
  begin
//    FLogName := ExtractFilePath (cheminDest) + 'GenerationPaie.log';
    FLogName := IncludeTrailingBackslash (TCBPPath.GetCegidDataDistri) + 'GenerationPaie.log';
    SetControlText('FICHIERLOG', FLogName);
  end;

  AssignFile(FLogFile, FLogName);
  {$I-}
  if FileExists(FLogName) then
  begin
    Append(FLogFile);
  end else
    ReWrite(FLogFile);
  {$I+}
  if IoResult<>0 then
  begin
   {$I-} CloseFile(FLogFile) ; {$I+}
   PGIBoxAF(TexteMessage[17]+ ' '+FLogName , ecran.caption);
   Exit;
  end;

  for i:=0 to FLogList.count-1 do
    Writeln(FLogFile, FLogList.Strings[i]);
  CloseFile(FLogFile) ;
end;

procedure TOF_BTPAIEGENER.Log(Msg: string);
begin
  FLogList.Add (' '+Msg);
end;

function RatioMesure(Cat, Mesure: string): Double;
var TOBM: TOB;
  X: Double;
begin
  TOBM := VH_GC.MTOBMEA.FindFirst(['GME_QUALIFMESURE', 'GME_MESURE'], [Cat, Mesure], False);
  X := 0;
  if TOBM <> nil then X := TOBM.GetValue('GME_QUOTITE');
  if X = 0 then X := 1.0;
  result := X;
end;

procedure TOF_BTPAIEGENER.TOBCopieChamp(FromTOB, ToTOB: TOB);
var i_pos, i_ind1: integer;
  FieldNameTo, FieldNameFrom, St: string;
  PrefixeTo, PrefixeFrom: string;
begin
  PrefixeFrom := TableToPrefixe(FromTOB.NomTable);
  PrefixeTo := TableToPrefixe(ToTOB.NomTable);
  for i_ind1 := 1 to FromTOB.NbChamps do
  begin
    FieldNameFrom := FromTOB.GetNomChamp(i_ind1);
    St := FieldNameFrom;
    i_pos := Pos('_', St);
    System.Delete(St, 1, i_pos - 1);
    FieldNameTo := PrefixeTo + St;
    if ToTOB.FieldExists(FieldNameTo) then
      ToTOB.PutValue(FieldNameTo, FromTOB.GetValue(FieldNameFrom))
  end;
end;

procedure TOF_BTPAIEGENER.GenereSection(Etab: string; TobActivite, TOBArticles, TOBTiers, TOBAffaires: TOB);
var TobStruc, TobA, TobT, TobAct, TOBPiece, TobAff, TOBL: TOB;
  Cpt, LCpt, Codeaxe: string;
  NumAxe, MaxAxe, i_ind: integer;

  function CreerSection(typeStruc: T_TypeStructAna): string;
  begin
    result := '';
    if TobStruc <> nil then
    begin
      TobStruc.Free;
      TobStruc := nil;
    end;
    if (Etab <> '') then TobStruc := ChargeAnaGescom(Codeaxe, Etab, 'PAI', typeStruc);
    if (TobStruc = nil) then TobStruc := ChargeAnaGescom(Codeaxe, '', 'PAI', typeStruc);
    if (TobStruc <> nil) then
      result := TraiteAnaGescom(CodeAxe, TobStruc, TOBPiece, TobL, TOBA, TobT, typeStruc);
  end;

begin
  if VH_GC.GCVENTAXE3 then MaxAxe := 3
  else if VH_GC.GCVENTAXE2 then MaxAxe := 2 else MaxAxe := 1;
  TOBPiece := TOB.Create('PIECE', nil, -1);
  TOBL := TOB.Create('LIGNE', nil, -1);
  TobStruc := nil;
  for i_ind := 0 to TobActivite.Detail.Count - 1 do
  begin
    TobAct := TobActivite.detail[i_ind];
    TOBCopieChamp(TobAct, TOBPiece);
    TOBCopieChamp(TobAct, TOBL);
    TobA := TOBArticles.FindFirst(['GA_ARTICLE'], [TobAct.GetValue('BCO_ARTICLE')], False);
    TobT := TOBTiers.FindFirst(['T_TIERS'], [TobAct.GetValue('AFF_TIERS')], False);
    TobAff := TOBAffaires.FindFirst(['AFF_AFFAIRE'], [TobAct.GetValue('BCO_AFFAIRE')], False);
    if trim(Etab) = '' then Etab := TobAff.Getvalue('AFF_ETABLISSEMENT');
    for NumAxe := 1 to MaxAxe do
    begin
      cpt := '';
      LCpt := '';
      Codeaxe := 'A' + IntToStr(NumAxe);
      Cpt := CreerSection(tCodesection);
      Cpt:=copy(Cpt,1,VH^.Cpta[AxeToFb(Codeaxe)].lg);
      if trim(Cpt) <> '' then
      begin
        TobAct.AddChampSupValeur('CODESECTION' + IntToStr(NumAxe), Cpt);
        LCpt := CreerSection(tLibSection);
        if trim(LCpt) <> '' then
          TobAct.AddChampSupValeur('LIBSECTION' + IntToStr(NumAxe), LCpt);
      end;
    end;
  end;
  if TobStruc <> nil then
  begin
    TobStruc.Free;
    TobStruc := nil;
  end;
  TOBPiece.free;
  TOBL.free;

end;

procedure TOF_BTPAIEGENER.BalanceLaGen;
var SL, Fichlard: TStringList;
  RubList: TRubList;
  NumsOrdre, AnaListVen: TAnaList;
  Rub: TRub;
  Q: TQuery;
  TobRes, TobDet, TobAktiv, TobActivMaj, TobMaj, TobD2, TOBArticles, TOBTiers, TOBAffaires: TOB;
  perdlo, perdhi, perdactlo, perdacthi, DateEntree: TDate;
  i, i_act, i_rub, i_ana, nt, anomalies, numorder, NumAxe, MaxAxe, NombreMHE, NbLigAct, CompteurPaieFic: integer;
  OldDS: Char;
  d, r: Double;
  Paramz, SQL, StReq, StMsg, FichierMvt, AccesFichierMvt, CurLigne: string;
  CSection, LSection, CExercice, LExercice, NoDossier: string;
  LiaisonY2 : boolean;
begin
  LiaisonY2 := GetParamSocSecur('SO_BTLIENPAIEY2',false);
  //
  if not GetParamSocSecur('SO_AFLIENPAIEANA',False) and not GetParamSocSecur('SO_AFLIENPAIEVAR',False) then
  begin
    PGIInfoAF(textemessage[3], ecran.caption);
    Exit;
  end;
  if not ExisteSQL ('SELECT ACP_RUBRIQUE FROM ACTIVITEPAIE') then
  begin
    PGIInfoAF(textemessage[19], ecran.caption);
    Exit;
  end;
  Paramz := AGLLanceFiche('BTP', 'BTPAIEGENERCHOIX', '', '', '');
  if Paramz = '' then exit;
  OldDS := DecimalSeparator;
  DecimalSeparator := '.';
  nt := 0;
  anomalies := 0;
  NbLigAct := 0;

  NumsOrdre := TAnaList.Create;
  TobRes := TOB.Create('Mom_RESSOURCE', nil, -1);
  SQL := 'SELECT ARS_RESSOURCE, ARS_SALARIE, ARS_LIBELLE, ARS_MENSUALISE,ARS_ETABLISSEMENT,PSA_DATEENTREE,PSA_PROFIL '+
    'FROM RESSOURCE LEFT JOIN SALARIES ON ARS_SALARIE=PSA_SALARIE';
  if TraiteEnregMulTable(TFMul(Ecran), SQL, 'ARS_RESSOURCE', 'RESSOURCE', 'ARS_RESSOURCE', 'RESSOURCE', TobRes, true) <> tteOK
  then
  begin
    TobRes.Free;
    exit;
  end;
  SL := TStringList.Create;
  SL.Text := StringReplace(Paramz, ';', Chr(VK_RETURN), [rfReplaceAll]);

  StMsg := SL.Values['EXERCICE'];
  CExercice := ReadTokenPipe(StMsg, ',');
  LExercice := StMsg;
  StMsg := SL.Values['PERIODE'];
  perdlo := StrToDate(ReadTokenPipe(StMsg, ','));
  perdhi := StrToDate(StMsg);
  StMsg := SL.Values['ACTIVITE'];
  perdactlo := StrToDate(ReadTokenPipe(StMsg, ','));
  perdacthi := StrToDate(StMsg);
  CheminDest := IncludeTrailingBackSlash(SL.Values['DEST']);

  if VH_GC.GCVENTAXE3 then MaxAxe := 3
  else if VH_GC.GCVENTAXE2 then MaxAxe := 2 else MaxAxe := 1;

  Fichlard := TStringList.Create;
  Fichlard.Add('***DEBUT***');
  FLogList := TStringList.Create;
  Log('');
  Log(TexteMessage[5] + DateTimeToStr(Now));
  Log(LExercice);
  Log(TexteMessage[6] + DateToStr(perdlo) + ' - ' + DateToStr(perdhi));
  Log(TexteMessage[7] + DateToStr(perdactlo) + ' - ' + DateToStr(perdacthi));
  Log('');


  // -- DOSSIER-SOCIETE
  Q := OpenSQL('SELECT PEX_DATEDEBUT, PEX_DATEFIN FROM EXERSOCIAL WHERE PEX_EXERCICE="' + CExercice + '"', true);
  if Q.EOF then
  begin
    Log(TexteMessage[9] + CExercice + ' - ' + LExercice);
    inc(anomalies);
  end;
  CurLigne := '000' + Sep;
  // récupération numéro de dossier si gestion multi-dossiers (ex : COTE)
  NoDossier := GetParamSocSecur ('SO_NODOSSIER', '000000');
  if NoDossier = '' then NoDossier:='000000';
  //
  CurLigne := CurLigne + NoDossier + Sep;
  CurLigne := CurLigne + FormatDateTime('dd"/"mm"/"yyyy', Q.FindField('PEX_DATEDEBUT').Value) + Sep + FormatDateTime('dd"/"mm"/"yyyy',
    Q.FindField('PEX_DATEFIN').Value);
  Ferme(Q);
  Fichlard.Add(CurLigne);
  TobActivMaj := TOB.Create('Maj activite', nil, -1);
  InitMove(TobRes.Detail.Count, '');
  for i := 0 to TobRes.Detail.Count - 1 do
  begin
    TobDet := TobRes.Detail[i];

    if (TobDet.GetValue('ARS_SALARIE') = null) or (TobDet.GetValue('ARS_SALARIE') = '') then
    begin
      StMsg := TobDet.GetValue('ARS_LIBELLE') + TexteMessage[10];
      Log(StMsg);
      inc(anomalies);
    end else if (GetParamSocSecur('SO_PGLIENRESSOURCE',False)) and
                (not ExisteSQL('SELECT PSA_SALARIE FROM SALARIES WHERE PSA_SALARIE="' + TobDet.GetValue('ARS_SALARIE') + '"')) then
    begin
      StMsg := TobDet.GetValue('ARS_LIBELLE') + Format(TexteMessage[11], [TobDet.GetValue('ARS_SALARIE')]);
      Log(StMsg);
      inc(anomalies);
    end else
    begin
      // -- Activité -> Rubriques
      TobAktiv := TOB.Create('', nil, -1);
      TOBArticles := TOB.Create('Les Articles', nil, -1);
      TOBTiers := TOB.Create('Les Tiers', nil, -1);
      TOBAffaires := TOB.Create('Les Affaires', nil, -1);
      RubList := TRubList.Create;
      AnaListVen := TAnaList.Create;
      try

        Q := OpenSQL('SELECT CONSOMMATIONS.*, GA_TYPEARTICLE, GA_FAMILLENIV1,GA_FAMILLENIV2,GA_FAMILLENIV3, AFF_TIERS ' +
          ' FROM CONSOMMATIONS LEFT JOIN ARTICLE ON BCO_ARTICLE=GA_ARTICLE' +
          ' LEFT JOIN AFFAIRE ON BCO_AFFAIRE=AFF_AFFAIRE' +
          ' WHERE BCO_RESSOURCE="' + TobDet.GetValue('ARS_RESSOURCE') + '" ' +
          ' AND BCO_DATEMOUV>="' + USDateTime(perdactlo) + '" AND BCO_DATEMOUV<="' + USDateTime(perdacthi) + '"' +
          ' AND BCO_PAIENUMFIC = 0 ', true);

        if not Q.EOF then TobAktiv.LoadDetailDB('CONSOMMATIONS', '', '', Q, false);
        Ferme(Q);
        if TobAktiv.detail.count = 0 then
        begin
          StMsg := TobDet.GetValue('ARS_LIBELLE') + TexteMessage[4];
          Log(StMsg);
          continue;
        end else
          NbLigAct := NbLigAct + TobAktiv.detail.count;

        Q := OpenSQL('SELECT * FROM ARTICLE WHERE GA_ARTICLE IN (SELECT BCO_ARTICLE FROM CONSOMMATIONS' +
          ' WHERE BCO_RESSOURCE="' + TobDet.GetValue('ARS_RESSOURCE') + '"' +
          ' AND BCO_DATEMOUV>="' + USDateTime(perdactlo) + '" AND BCO_DATEMOUV<="' + USDateTime(perdacthi) + '" )', true);
        if not Q.EOF then TOBArticles.LoadDetailDB('ARTICLE', '', '', Q, false, false);
        Ferme(Q);

        Q := OpenSQL('SELECT * FROM TIERS WHERE T_TIERS IN (SELECT AFF_TIERS FROM AFFAIRE LEFT JOIN CONSOMMATIONS ON AFF_AFFAIRE=BCO_AFFAIRE' +
          ' WHERE BCO_RESSOURCE="' + TobDet.GetValue('ARS_RESSOURCE') + '" ' +
          ' AND BCO_DATEMOUV>="' + USDateTime(perdactlo) + '" AND BCO_DATEMOUV<="' + USDateTime(perdacthi) + '" )', true);
        if not Q.EOF then TOBTiers.LoadDetailDB('TIERS', '', '', Q, false, false);
        Ferme(Q);

        Q := OpenSQL('SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE IN (SELECT BCO_AFFAIRE FROM CONSOMMATIONS' +
          ' WHERE BCO_RESSOURCE="' + TobDet.GetValue('ARS_RESSOURCE') + '"' +
          ' AND BCO_DATEMOUV>="' + USDateTime(perdactlo) + '" AND BCO_DATEMOUV<="' + USDateTime(perdacthi) + '" )', true);
        if not Q.EOF then TOBAffaires.LoadDetailDB('AFFAIRE', '', '', Q, false, false);
        Ferme(Q);

        if TobAktiv.Detail.Count > 0 then
        begin
          Log(format(TexteMessage[12], [TobDet.GetValue('ARS_LIBELLE')]) + TobDet.GetValue('ARS_SALARIE'));
          if (GetParamSocSecur('SO_AFLIENPAIEANA',False)) and (GetParamSocSecur('SO_PGANALYTIQUE',False)) then
            GenereSection('', TobAktiv, TOBArticles, TOBTiers, TOBAffaires);
        end;
        // 1ère passe, totalisage par rubrique
        for i_act := TobAktiv.Detail.Count - 1 downto 0 do
        begin
          TobD2 := TobAktiv.Detail[i_act];
          StReq := 'SELECT ACP_RUBRIQUE,ACP_RESSOURCE,ACP_TYPEHEURE,ACP_TYPEARTICLE,ACP_ARTICLE,ACP_PROFIL,ACP_VENTANARUB,' +
            'ACP_FAMILLENIV1,ACP_FAMILLENIV2,ACP_FAMILLENIV3,ACP_MENSUALISE,'+
            'PRM_TYPEBASE,PRM_TYPEMONTANT,PRM_BASEMTQTE ' +
            ' FROM ACTIVITEPAIE LEFT JOIN REMUNERATION ON ACP_RUBRIQUE=PRM_RUBRIQUE' +
            ' WHERE ';
            if nodossier <> '' then
            begin
            	StReq := StReq + '(';
            end;
            stReq := StREq + 'ACP_RESSOURCE="' + TobD2.GetValue('BCO_RESSOURCE') + '"' +
            ' OR ACP_TYPEARTICLE="' + TobD2.GetValue('GA_TYPEARTICLE') + '"' +
            ' OR ACP_ARTICLE="' + TobD2.GetValue('BCO_ARTICLE') + '"';
          if TobDet.GetValue('PSA_PROFIL') <> null then
            StReq := StReq +  ' OR ACP_PROFIL="' + TobDet.GetValue('PSA_PROFIL') + '"' ;
          if TobD2.GetValue('BCO_TYPEHEURE') <> '' then
            StReq := StReq + ' OR ACP_TYPEHEURE="' + TobD2.GetValue('BCO_TYPEHEURE') + '"';
          if TobD2.GetValue('GA_FAMILLENIV1') <> '' then
            StReq := StReq + ' OR ACP_FAMILLENIV1="' + TobD2.GetValue('GA_FAMILLENIV1') + '"';
          if TobD2.GetValue('GA_FAMILLENIV2') <> '' then
            StReq := StReq + ' OR ACP_FAMILLENIV2="' + TobD2.GetValue('GA_FAMILLENIV2') + '"';
          if TobD2.GetValue('GA_FAMILLENIV3') <> '' then
            StReq := StReq + ' OR ACP_FAMILLENIV3="' + TobD2.GetValue('GA_FAMILLENIV3') + '"';

         // utiliser nodossier
         // pour les rubriques se terminant par 5,7,9 prendre nodossier pour les autres 000000
          if nodossier <> '' then
          begin
          	stReq := stReq + ') AND ((RIGHT(ACP_RUBRIQUE,1) IN ("5","7","9") AND PRM_NODOSSIER="'+NODOSSIER+'") OR ((RIGHT(ACP_RUBRIQUE,1) NOT IN ("5","7","9") AND PRM_NODOSSIER="000000")))';
          end;

          StReq := StReq + ' ORDER BY ACP_TYPEHEURE, ACP_TYPEARTICLE DESC ';

          Q := OpenSQL(StReq, true);
          TobD2.LoadDetailDB('ACTIVITEPAIE', '', '', Q, false);
          Ferme(Q);
          if TobD2.Detail.Count = 0 then
          begin //Log('Aucune rubrique correspondante');
            TobAktiv.Detail[i_act].Free; // Pas besoin : on vire
            Continue;
          end;

          for i_rub := TobD2.Detail.Count - 1 downto 0 do
          begin
            if ((trim(TobD2.Detail[i_rub].GetValue('ACP_ARTICLE')) <> '')
              and (TobD2.GetValue('BCO_ARTICLE') <> TobD2.Detail[i_rub].GetValue('ACP_ARTICLE'))) then
            begin
              TobD2.Detail[i_rub].Free;
              Continue;
            end;
            if ((trim(TobD2.Detail[i_rub].GetValue('ACP_RESSOURCE')) <> '')
              and (TobD2.GetValue('BCO_RESSOURCE') <> TobD2.Detail[i_rub].GetValue('ACP_RESSOURCE'))) then
            begin
              TobD2.Detail[i_rub].Free;
              Continue;
            end;
            if ((trim(TobD2.Detail[i_rub].GetValue('ACP_PROFIL')) <> '')
              and (TobDet.GetValue('PSA_PROFIL') <> TobD2.Detail[i_rub].GetValue('ACP_PROFIL'))) then
            begin
              TobD2.Detail[i_rub].Free;
              Continue;
            end;
            if ((trim(TobD2.Detail[i_rub].GetValue('ACP_TYPEARTICLE')) <> '')
              and (TobD2.GetValue('GA_TYPEARTICLE') <> TobD2.Detail[i_rub].GetValue('ACP_TYPEARTICLE'))) then
            begin
              TobD2.Detail[i_rub].Free;
              Continue;
            end;
            if (TobD2.GetValue('BCO_TYPEHEURE') <> TobD2.Detail[i_rub].GetValue('ACP_TYPEHEURE')) then
            begin
              TobD2.Detail[i_rub].Free;
              Continue;
            end;
            if (TobDet.GetValue('ARS_MENSUALISE') = 'X') and (TobD2.Detail[i_rub].GetValue('ACP_MENSUALISE') = '-') then
            begin
              TobD2.Detail[i_rub].Free;
              Continue;
            end;
            if ((trim(TobD2.GetValue('GA_FAMILLENIV1')) <> '') and (trim(TobD2.Detail[i_rub].GetValue('ACP_FAMILLENIV1')) <> '')
              and (TobD2.GetValue('GA_FAMILLENIV1') <> TobD2.Detail[i_rub].GetValue('ACP_FAMILLENIV1'))) then
            begin
              TobD2.Detail[i_rub].Free;
              Continue;
            end;
            if ((trim(TobD2.GetValue('GA_FAMILLENIV2')) <> '') and (trim(TobD2.Detail[i_rub].GetValue('ACP_FAMILLENIV2')) <> '')
              and (TobD2.GetValue('GA_FAMILLENIV2') <> TobD2.Detail[i_rub].GetValue('ACP_FAMILLENIV2'))) then
            begin
              TobD2.Detail[i_rub].Free;
              Continue;
            end;
            if ((trim(TobD2.GetValue('GA_FAMILLENIV3')) <> '') and (trim(TobD2.Detail[i_rub].GetValue('ACP_FAMILLENIV3')) <> '')
              and (TobD2.GetValue('GA_FAMILLENIV3') <> TobD2.Detail[i_rub].GetValue('ACP_FAMILLENIV3'))) then
            begin
              TobD2.Detail[i_rub].Free;
              Continue;
            end;
          end;
          for i_rub := 0 to TobD2.Detail.Count - 1 do
            with RubList.Rubrik[TobD2.Detail[i_rub].GetValue('ACP_RUBRIQUE')] do
            begin
              Libelle := TobD2.GetValue('BCO_LIBELLE');
              d := 0.0;
              if (TobD2.Detail[i_rub].GetValue('PRM_TYPEBASE') = '00')
                or (TobD2.Detail[i_rub].GetValue('PRM_TYPEBASE') = '03') then // '00' = Variable
              begin
                d := TobD2.GetValue('BCO_QUANTITE');

                // Conversion ici avec MEA en fonction de l'unité
                // et de l'unité de la rubrik paie -> PRM_BASEMTQTE 002=Jour 003=Heure
                if TobD2.Detail[i_rub].GetValue('PRM_BASEMTQTE') = '002' then r := 8 else r := 1;
                d := (d * RatioMesure('TEM', TobD2.GetValue('BCO_QUALIFQTEMOUV'))) / r;

                Qte := Qte + d;
                GotQte := true;
              end;
              if (TobD2.Detail[i_rub].GetValue('PRM_TYPEMONTANT') = '00')
                or (TobD2.Detail[i_rub].GetValue('PRM_TYPEMONTANT') = '03') then
              begin
                Montant := Montant + TobD2.GetValue('BCO_MONTANTPR');
                GotMontant := true;
              end;
              Ventiled := (TobD2.Detail[i_rub].GetValue('ACP_VENTANARUB') = 'X');
              if Ventiled then
              begin
                // Faire un Select AU PREALABLE pour chopper tous les axes et sections définis
                for NumAxe := 1 to MaxAxe do
                  if TobD2.FieldExists('CODESECTION' + IntToStr(NumAxe)) then
                  begin
                    CSection := TobD2.GetValue('CODESECTION' + IntToStr(NumAxe));
                    LSection := TobD2.GetValue('LIBSECTION' + IntToStr(NumAxe));
                    with VentAnal.Section['A' + IntToStr(NumAxe), CSection, LSection] do
                      if GotQte then Value := Value + d
                      else Value := Value + TobD2.GetValue('BCO_MONTANTPR');
                  end;
              end;
            end;
          // Ventilation analytique par salarié
          if (GetParamSocSecur('SO_AFLIENPAIEANA',False)) and (GetParamSocSecur('SO_PGANALYTIQUE',False)) and (TobD2.GetValue('BCO_NATUREMOUV') = 'MO') and
          	 (TobD2.Detail.Count > 0) then
          begin
            if (TobD2.Detail[0] <> Nil) then
            begin
               if TobD2.Detail[0].GetValue('ACP_VENTANARUB') = '-' then
               begin
                 for NumAxe := 1 to MaxAxe do
                   if TobD2.FieldExists('CODESECTION' + IntToStr(NumAxe)) then
                   begin
                     CSection := TobD2.GetValue('CODESECTION' + IntToStr(NumAxe));
                     LSection := TobD2.GetValue('LIBSECTION' + IntToStr(NumAxe));
                     with AnaListVen.Section['A' + IntToStr(NumAxe), CSection, LSection] do
                       Value := Value + TobD2.GetValue('BCO_QUANTITE');
                   end;
               end;
            end;
          end;

          // Toper les lignes de consos avec le numéro de fichier
          TobMaj := Tob.Create('CONSOMMATIONS', TobActivMaj, -1);
          TobMaj.Dupliquer(TobD2, false, true, false);
        end;

        // Modif BRL 8/10/2012 : prendre date d'entrée en date de début de période si celle-ci est plus récente.
        DateEntree := StrToDate(tobdet.getvalue('PSA_DATEENTREE'));
        if  DateEntree > perdlo then
          SL.Values['PERIODE']:=DateToStr(DateEntree)+','+DateToStr(perdhi)
        else
          SL.Values['PERIODE']:=DateToStr(perdlo)+','+DateToStr(perdhi);

        for i_rub := 0 to RubList.Count - 1 do
        begin
          Rub := RubList.Rub[i_rub];
          // -- LIGNE IMPORT TYPE RUBRIQUE
          if GetParamSocSecur('SO_AFLIENPAIEVAR',False) then
          begin
            CurLigne := 'MHE' + Sep +
              TobDet.GetValue('ARS_SALARIE') + Sep + // Code Salarié
            StringReplace(SL.Values['PERIODE'], ',', Sep, []) + Sep + // Dates Début/Fin Période
            inttostr(NumsOrdre.Section[TobDet.GetValue('ARS_SALARIE') + SL.Values['PERIODE'] + Rub.CodeRub, '', ''].NumVen) + Sep +
            Rub.CodeRub + Sep + Sep;
// modif BRL 11/02 : plus de libellé => celui de la rubrique sera repris par défaut en paie
//            Rub.CodeRub + Sep + Rub.Libelle + Sep;

            if (Rub.GotQte) and (Rub.GotMontant) then CurLigne := CurLigne + 'BM' + Sep else
              if not Rub.GotMontant then CurLigne := CurLigne + 'BAS' + Sep else
              if not Rub.GotQte then CurLigne := CurLigne + 'MON' + Sep;

            if Rub.GotQte then CurLigne := CurLigne + Format('%g', [Rub.Qte]);
            CurLigne := CurLigne + Sep + Sep + Sep;
            if Rub.GotMontant then CurLigne := CurLigne + Format('%g', [Rub.Montant]);

            Fichlard.Add(CurLigne);

            if (GetParamSocSecur('SO_AFLIENPAIEANA',False)) and (GetParamSocSecur('SO_PGANALYTIQUE',False)) and Rub.Ventiled then // Ventil spécifique
            begin
//              Rub.PourcentRubrique;
		          Rub.PourcentRubrique('A1');
    		      if MaxAxe > 1 then
        		  	Rub.PourcentRubrique('A2');
          		if MaxAxe > 2 then
          			Rub.PourcentRubrique('A3');

              for i_ana := 0 to Rub.VentAnal.Count - 1 do // Parcourage des ventilations
              begin
          	    CreerSectionVolee (Rub.VentAnal.Ana[i_ana].Section,Rub.VentAnal.Ana[i_ana].Libelle,StrToInt(copy(Rub.VentAnal.Ana[i_ana].Axe,2,1)));
                if LiaisonY2 then
                begin
                  CurLigne := 'ANA' + Sep +
                              TobDet.GetValue('ARS_SALARIE') + Sep + // Code Salarié
                              StringReplace(SL.Values['PERIODE'], ',', Sep, []) + Sep + // Dates Début/Fin Période
                              Rub.CodeRub + Sep +
                              Rub.VentAnal.Ana[i_ana].Axe + Sep + // Mettre ici l'axe analytique
                              'S1' + Sep + // Mettre ici le sous plan
                              Rub.VentAnal.Ana[i_ana].Section + Sep + // Mettre ici la section analytique
                              Format('%.2f', [Rub.VentAnal.Ana[i_ana].Pourcent]) + Sep +
                              inttostr(Rub.VentAnal.Ana[i_ana].NumVen) + Sep +
                              Rub.VentAnal.Ana[i_ana].Libelle; //Libelle de la section rajoutée
                end else
                begin
                  CurLigne := 'ANA' + Sep +
                              TobDet.GetValue('ARS_SALARIE') + Sep + // Code Salarié
                              StringReplace(SL.Values['PERIODE'], ',', Sep, []) + Sep + // Dates Début/Fin Période
                              Rub.CodeRub + Sep +
                              Rub.VentAnal.Ana[i_ana].Axe + Sep + // Mettre ici l'axe analytique
                              Rub.VentAnal.Ana[i_ana].Section + Sep + // Mettre ici la section analytique
                              Format('%.2f', [Rub.VentAnal.Ana[i_ana].Pourcent]) + Sep +
                              inttostr(Rub.VentAnal.Ana[i_ana].NumVen) + Sep +
                              Rub.VentAnal.Ana[i_ana].Libelle; //Libelle de la section rajoutée
                end;
                Fichlard.Add(CurLigne);
              end;
            end;
          end;
        end;
        if (GetParamSocSecur('SO_AFLIENPAIEANA',False)) and (GetParamSocSecur('SO_PGANALYTIQUE',False)) then
        begin
          AnaListVen.Pourcentize('A1');
          if MaxAxe > 1 then
          	AnaListVen.Pourcentize('A2');
          if MaxAxe > 2 then
          	AnaListVen.Pourcentize('A3');
          for i_ana := 0 to AnaListVen.Count - 1 do
          begin
          	CreerSectionVolee (AnaListVen.Ana[i_ana].Section,AnaListVen.Ana[i_ana].Libelle,StrToInt(copy(AnaListVen.Ana[i_ana].Axe,2,1)));
            if LiaisonY2 then
            begin
              CurLigne := 'VSA' + Sep +
                          TobDet.GetValue('ARS_SALARIE') + Sep + // Code Salarié
                          StringReplace(SL.Values['PERIODE'], ',', Sep, []) + Sep + // Dates Début/Fin Période
                          AnaListVen.Ana[i_ana].Axe + Sep + // Mettre ici l'axe analytique
                          'S1' + Sep + // Mettre ici l'axe analytique
                          AnaListVen.Ana[i_ana].Section + Sep + // Mettre ici la section analytique
                          Format('%.2f', [AnaListVen.Ana[i_ana].Pourcent]) + Sep +
                          inttostr(AnaListVen.Ana[i_ana].NumVen) + Sep +
                          AnaListVen.Ana[i_ana].Libelle; //Libelle de la section rajoutée
            end else
            begin
              CurLigne := 'VSA' + Sep +
                          TobDet.GetValue('ARS_SALARIE') + Sep + // Code Salarié
                          StringReplace(SL.Values['PERIODE'], ',', Sep, []) + Sep + // Dates Début/Fin Période
                          AnaListVen.Ana[i_ana].Axe + Sep + // Mettre ici l'axe analytique
                          AnaListVen.Ana[i_ana].Section + Sep + // Mettre ici la section analytique
                          Format('%.2f', [AnaListVen.Ana[i_ana].Pourcent]) + Sep +
                          inttostr(AnaListVen.Ana[i_ana].NumVen) + Sep +
                          AnaListVen.Ana[i_ana].Libelle; //Libelle de la section rajoutée
            end;
            Fichlard.Add(CurLigne);
          end;
        end;
      finally
        RubList.Free;
        AnaListVen.Free;
        TobAktiv.Free;
        TOBArticles.free;
        TOBTiers.free;
        TOBAffaires.free;
      end;

      inc(nt);
    end;
    MoveCur(False);
  end;
  FiniMove;
  DecimalSeparator := OldDS;
  Fichlard.Add('***FIN***');

  if (TobActivMaj.detail.count > 0)then
  begin
    CompteurPaieFic:=GetParamSocSecur('SO_AFLIENPAIEFIC',0);
  	numorder := CompteurPaieFic + 1;
    SetParamSoc('SO_AFLIENPAIEFIC', numorder);
    FichierMvt := 'BTPAIE';
    FichierMvt := FichierMvt + Format('%.6d', [numorder]) + FormatDateTime('yyyymmddhhmm', Now) + '.TXT';
    AccesFichierMvt := IncludeTrailingBackSlash(SL.Values['DEST']) + FichierMvt;
    Fichlard.SaveToFile(AccesFichierMvt);
    Log(TexteMessage[8] + Format('%.6d', [numorder]) + ' ' + AccesFichierMvt);

    for i_act:=0 to TobActivMaj.detail.count-1 do
      TobActivMaj.detail[i_act].Putvalue('BCO_PAIENUMFIC', numorder);
    if not TobActivMaj.UpdateDB(true) then
      Log(TexteMessage[13]);
    StMsg := format(TexteMessage[14], [inttostr(nt)]) + inttostr(TobRes.Detail.Count);
    Log(StMsg);
  end else
  begin
    StMsg := TexteMessage[16] + DateTostr(perdactlo) + ' - ' + DateTostr(perdacthi);
    Log(StMsg);
  end;
  Log('');

  EcrireFichierLog;
  if anomalies > 0 then
    StMsg := StMsg + Chr(VK_RETURN) + inttostr(anomalies) + TexteMessage[15] + FLogName;
  // Confirmation de l'envoie dans la paie - Traitement dans UtofPgImportFic
  if not GetParamSocSecur('SO_AFLIENPAIEDEC',False) and (anomalies = 0) and (nt > 0) and (TobActivMaj.detail.count > 0) then
    AGLLanceFiche('BTP', 'BTPAIEGENER', '', '', AccesFichierMvt + ';AUTO')
  else PGIInfoAF(StMsg, ecran.caption);


  NumsOrdre.Free;
  Fichlard.Free;
  FLogList.Free;
  SL.Free;
  TobRes.Free;
end;

procedure TOF_BTPAIEGENER.Gen;
var F: TFMul;
begin
  F := TFMul(Ecran);
  if (F.FListe.NbSelected = 0) and (not F.FListe.AllSelected)
    then PGIInfoAF(textemessage[18], ecran.caption)
  else BalanceLaGen;
end;

procedure BTLanceFiche_ActivGenerPaie;
begin
  AGLLanceFiche ('BTP','BTPAIEGENER_MUL','','','');
end;


/////////////// Procedure appeleé par le bouton Validation //////////////

procedure AGLGenereFichierMouvements(Parms: array of variant; nb: integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(longint(Parms[0]));
  if (F is TFMul) then TOTOF := TFMul(F).LaTOF else exit;
  if (TOTOF is TOF_BTPAIEGENER) then TOF_BTPAIEGENER(TOTOF).Gen else exit;
end;

initialization
  RegisterClasses([TOF_BTPAIEGENER]);
  RegisterAglProc('GenereFichierMouvements', true, 0, AGLGenereFichierMouvements);
end.
