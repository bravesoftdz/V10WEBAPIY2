{***********UNITE*************************************************
Auteur  ...... : M.N. GARNIER
Créé le ...... : 23/12/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
unit UtilAlertes;

interface
uses
  Variants,

{$IFDEF EAGLCLIENT}
  MaineAgl, eFiche, eFichList, eFichGrid
{$ELSE}
{$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, db, {$ENDIF}FE_Main, Fiche, FichList, FichGrid
{$ENDIF}

  , M3FP, Utob, SysUtils, HMsgBox, HEnt1, HCtrls, ExtCtrls, ComCtrls, windows, HRichOle,
  UTOF, stdctrls, AGLInit, Classes, Controls, Forms, UTom, UtilPgi,
  SaisieList
  ;
type
  The_Alerte = class
    Form: TForm;
    Prefixe: string;
    Evenement: string;
    TobTable: Tob;
    TobsInitiales: Tob;
    FOM: TOM;
  private
    function GetOM: TOM;
  public
    constructor create;
    class function ExecuteAlerte(const F: TForm; const FPrefixe, FEVent: string; const FTob, FTobsInit: TOB; const A_OM: TOM = nil): Boolean;
    // --------
    property OM: TOM read GetOM write FOM;
  end;

function Rt_VerifierAlertes(ObjAlerte: The_Alerte): boolean;
function Rt_VerifCondi(TobAlertes, TobTable: tob): boolean;
function Rt_TraducCondi(TobTable: tob; champ, operateur, valchamp: string): boolean;
function Rt_NotifieAlerte(TobAlertes, TobTable: Tob): boolean;
function Rt_isFieldModified(OM: TOM; stChamp: string; TobApres, TobAvant: tob): boolean;
function ExecuteAlerteUpdate(F: TForm; DoFree: boolean): boolean;
function ExecuteAlerteLoad(F: TForm; DoFree: boolean): boolean;
function ExecuteAlerteDelete(F: TForm; DoFree: boolean): boolean;
procedure DoCreateTobFinale(F: TForm);
procedure DoCreateTobOrigine(F: TForm);
procedure DoFreeTobFinale(F: TForm);
procedure DoFreeTobOrigine(F: TForm);
function ChargeTobAlerte(ObjAlerte: The_Alerte; bAlerteCroisee, bAlerte7: boolean): tob;
function AlerteActive(prefixe: string): boolean;
function ExisteAlerteActive: boolean;
function GetStateFicheAlerte(F: TForm): variant;
procedure Rt_LogAlerte(TobAlertDeclench: Tob);
function RecupCodeEvent(TobAlertes: tob): string;
Function RT_DescroAlerte(TobTable : tob) : string;

var
  FMonAlerte: The_Alerte;

implementation

uses
  WCommuns, YAlertesConst, EntPgi, YTableAlertes_tom
{$IFDEF GPAO}
  , EntGP
{$ENDIF GPAO}
  ,BTPUtil,HRichEdt
  ,CbpMCD

  ;


function localRichToString(const Msg: HString): HString;
var
  Memo: THRichEdit;
  Panel: TPanel;
begin
  Panel := TPanel.Create(nil);
  try
    Panel.Visible := False;
    Panel.ParentWindow := GetDesktopWindow;
    Memo := THRichEdit.Create(Panel);
    Memo.Parent := Panel;
    Memo.Visible := False;
    Memo.Text := Msg; //Memo.Lines.Text := Msg;
    Result := ExRichToString(Memo);
    Memo.Free;
  finally
    Panel.Free;
  end
end;

{***********A.G.L.***********************************************
Auteur  ...... : M.N. GARNIER
Créé le ...... : 23/12/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function Rt_VerifierAlertes(ObjAlerte: The_Alerte): boolean;
var stEvent, ListeEvent, oldCode: string;
  TobAlertes: Tob;
  i, iTob: integer;
  existAlCh, tousModifie, bAlerteCroisee, bVerifCroisee, bAlerte7: boolean;
  OM: TOM;

  { Notification sans écran : Flag dans la TobAlertes afin de pouvoir ramener le message
    d'alerte dans cette même tob, depuis la fonction Rt_VerifCondi }
  procedure FlagAlertSansEcran(const TobIdx: Integer);
    function isContextClickSingleDelete: Boolean;
    begin
      Result := Assigned(ObjAlerte.TobTable.Parent) and (ObjAlerte.TobTable.Parent.Detail.Count = 1) //-> Suppression mono enregistrement
                {$IFDEF GPAO}
                  and VH_GP.GetGlobalArgumentBoolean(GLOBARG_FROMBDELETE, True)                      //-> On provient du BSUPPRIMEClick de la wTof
                {$ENDIF GPAO}
    end;
  var
    T: Tob;
  begin
    T := TobAlertes.Detail[TobIdx];
    if V_Pgi.SilentMode or (not Assigned(ObjAlerte.Form) and not isContextClickSingleDelete()) then
      T.AddChampSupValeur(ALERT_SansEcran, wTrue);
  end;

begin
  result := true;
  if ObjAlerte.evenement = '' then exit;
  { cas particulier des alertes croisées tiers et articles }
  ListeEvent := ObjAlerte.evenement;
  stEvent := ReadToKenSt(ListeEvent);
  bAlerteCroisee := (stEvent[1] = '6') or (stEvent[1] = '7');
  bAlerte7 := (stEvent[1] = '7');

  TobAlertes := ChargeTobAlerte(ObjAlerte, bAlerteCroisee, bAlerte7);
  try
    if TobAlertes.detail.Count = 0 then exit;

    { on traite à part les alertes croisées }
    if bAlerteCroisee then
    begin
      bVerifCroisee := true;
      oldCode := '';
      { 1 er test : alertes conditionnées sur valeurs champs }
      for i := 0 to Pred(TobAlertes.detail.count) do
      begin
        if oldCode = '' then oldCode := TobAlertes.detail[i].GetString('YAL_ALERTE');

        if (TobAlertes.detail[i].GetString('YAL_ALERTE') <> oldCode) then
        begin
          if bVerifCroisee then
          begin
            FlagAlertSansEcran(i - 1);

            if not Rt_NotifieAlerte(TobAlertes.detail[i - 1], ObjAlerte.TobTable) then
            begin
              result := false;
              exit;
            end;
          end
          else
            bVerifCroisee := true;
        end;
        oldCode := TobAlertes.detail[i].GetString('YAL_ALERTE');

        if (not Rt_VerifCondi(TobAlertes.detail[i], ObjAlerte.TobTable)) then
          bVerifCroisee := false;
      end;
      if bVerifCroisee then
      begin
        FlagAlertSansEcran(TobAlertes.Detail.Count - 1);

        if not Rt_NotifieAlerte(TobAlertes.detail[TobAlertes.detail.count - 1], ObjAlerte.TobTable) then
        begin
          result := false;
          exit;
        end;
      end
    end
    else
    begin
      { 1 er test : alertes conditionnées sur conditions valeurs champs (pas evénements 4 et 5) }
      for i := 0 to Pred(TobAlertes.detail.count) do
      begin
        if (TobAlertes.detail[i].GetBoolean('YAL_EVENEMENT4')) or
          (TobAlertes.detail[i].GetBoolean('YAL_EVENEMENT5')) or
          (not Rt_VerifCondi(TobAlertes.detail[i], ObjAlerte.TobTable)) then continue;

        FlagAlertSansEcran(i);

        if not Rt_NotifieAlerte(TobAlertes.detail[i], ObjAlerte.TobTable) then
        begin
          result := false;
          exit;
        end;
      end;
      { 2 ème test : alertes conditionnées sur modifs de champs ou date anniversaire }
      existAlCh := false;
      tousModifie := true;
      oldCode := '';
      if Assigned(ObjAlerte.Form) or Assigned(ObjAlerte.FOM) then
      begin
        iTob := 0;
        OM := ObjAlerte.OM;
        if Assigned(OM) and (TobAlertes.detail.count > 0) then
          for i := 0 to TobAlertes.detail.count - 1 do
          begin
            { rupture sur code alerte : on teste si déclenchement }
            if oldCode = '' then oldCode := TobAlertes.detail[i].GetString('YAL_ALERTE');
            if (TobAlertes.detail[i].GetString('YAL_ALERTE') <> oldCode) then
            begin
              if (existAlCh = true) and (tousModifie = true) then
              begin
                FlagAlertSansEcran(iTob);

                if not Rt_NotifieAlerte(TobAlertes.detail[iTob], ObjAlerte.TobTable) then
                begin
                  result := false;
                  exit;
                end;
              end;
              tousModifie := true;
              existAlCh := false;
              oldCode := TobAlertes.detail[i].GetString('YAL_ALERTE');
            end;
            if (TobAlertes.detail[i].GetString('YAU_CHAMP') <> '') and
              (TobAlertes.detail[i].GetString('YAU_CHAMP') <> NULL) and
              (TobAlertes.detail[i].GetBoolean('YAL_EVENEMENT4')) and
              (TobAlertes.detail[i].GetString('YAU_TYPE') = 'C') then
            begin
              existAlCh := true;
              iTob := i;
              if not Rt_isFieldModified(OM, TobAlertes.detail[i].GetString('YAU_CHAMP'), ObjAlerte.TobTable, ObjAlerte.TobsInitiales) then
              begin
                tousModifie := false;
              end;
            end;
            if (TobAlertes.detail[i].GetString('YAU_CHAMP') <> '') and
              (TobAlertes.detail[i].GetString('YAU_CHAMP') <> NULL) and
              (TobAlertes.detail[i].GetBoolean('YAL_EVENEMENT5')) and
              (TobAlertes.detail[i].GetString('YAU_TYPE') = 'D') then
            begin
              if not ObjAlerte.TobTable.FieldExists(TobAlertes.detail[i].GetString('YAU_CHAMP')) then
              begin
                tousModifie := false;
              end;
              existAlCh := true;
              iTob := i;
               { test= (date du jour <= date champ anniv) et (date du jour + 15 >= date champ anniv) }
              if not (Date <= ObjAlerte.TobTable.GetDateTime(TobAlertes.detail[i].GetString('YAU_CHAMP'))) or
                not (PlusDate(Date, TobAlertes.detail[i].GetInteger('YAL_NBJOURS'), 'J') >=
                ObjAlerte.TobTable.GetDateTime(TobAlertes.detail[i].GetString('YAU_CHAMP'))) then
              begin
                tousModifie := false;
              end;
            end;
          end;
        if (existAlCh = true) and (tousModifie = true) then
        begin
          FlagAlertSansEcran(iTob);

          if not Rt_NotifieAlerte(TobAlertes.detail[iTob], ObjAlerte.TobTable) then
          begin
            result := false;
            exit;
          end;
        end
      end;
    end;
  finally
    { Ramène dans la Tob de la Tom métier, le(s) message(s) d'alerte(s)
      notifié(s) en cas d'utilisation du code métier sans écran }
    for i := 0 to Pred(TobAlertes.Detail.Count) do
      if TobAlertes.Detail[i].FieldExists(ALERT_TobMsg) then
      begin
        if not ObjAlerte.TobTable.FieldExists(ALERT_TobMsg) then
          ObjAlerte.TobTable.AddChampSup(ALERT_TobMsg, False);
        ObjAlerte.TobTable.SetString(ALERT_TobMsg, localRichToString(ObjAlerte.TobTable.GetString(ALERT_TobMsg))
                                                 + localRichToString(TobAlertes.Detail[i].GetString(ALERT_TobMsg)) + #13);
      end;

    FreeAndNil(TobAlertes);
  end
end;

{ si champ de la table de la TOM, OM.isFieldModified, sinon, comparaison avec
  tobInitiales, ceci pour Tierscompl et les tables infos compl }

function Rt_isFieldModified(OM: TOM; stChamp: string; TobApres, TobAvant: tob): boolean;
begin
  result := false;
  if pos('_', stChamp) = 0 then exit;
  if copy(stChamp, 1, pos('_', stChamp) - 1) = TableToPrefixe(OM.TableName) then
  begin
    result := OM.isFieldModified(stChamp);
    exit;
  end;
  result := (TobAvant.GetValue(stChamp) <> TobApres.GetValue(stChamp));
end;
{***********A.G.L.***********************************************
Auteur  ...... : M.N. GARNIER
Créé le ...... : 23/12/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function Rt_VerifCondi(TobAlertes, TobTable: tob): boolean;
var
  i: integer;
  Condi: array[1..3] of boolean;
  Lien: array[1..2] of string;
begin
  if (TobAlertes.GetString('YAD_CHAMP1') = '') or
    (TobAlertes.GetString('YAD_CHAMP1') = NULL) then
  begin
    result := true;
    exit;
  end;
  for i := 1 to 3 do
  begin
    Condi[i] := true;
    if i < 3 then Lien[i] := 'Et';
    if TobAlertes.GetString('YAD_CHAMP' + IntToStr(i)) <> '' then
    begin
      Condi[i] := Rt_TraducCondi(TobTable, TobAlertes.GetString('YAD_CHAMP' + IntToStr(i)), TobAlertes.GetValue('YAD_OPER' + IntToStr(i)),
        TobAlertes.GetValue('YAD_VAL' + IntToStr(i)));
      if i < 3 then
        if Trim(TobAlertes.GetValue('YAD_LIEN' + IntToStr(i))) <> '' then
          Lien[i] := TobAlertes.GetValue('YAD_LIEN' + IntToStr(i));
    end;
  end;
  if lien[1] = 'Et' then
  begin
    if lien[2] = 'Et' then
    begin
      if (Condi[1]) and (Condi[2]) and (Condi[3]) then Result := true
      else result := false;
    end
    else
    begin
      if Condi[1] and (Condi[2] or Condi[3]) then Result := true
      else result := false;
    end
  end
  else
    if lien[2] = 'Et' then
    begin
      if (Condi[1] or Condi[2]) and Condi[3] then Result := true
      else result := false;
    end
    else
    begin
      if Condi[1] or Condi[2] or Condi[3] then Result := true
      else result := false;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : M.N. GARNIER
Créé le ...... : 23/12/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function Rt_TraducCondi(TobTable: tob; champ, operateur, valchamp: string): boolean;
var
  val, val1, val2: string;
  Valeur1,Valeur2: Variant;
  TypeField: Char;
begin
  TypeField := wGetSimpleTypeField(champ);
  Valeur1:='';
  Valeur2:='';
  case TypeField of
    'N', 'I':
          begin
            Valeur1 := Valeur(TobTable.getString(champ));
            Valeur2 := Valeur(valchamp);
          end;
    'D':  begin
            if not IsDateVide(TobTable.getString(champ)) then
              valeur1 := FormatDateTime('yyyymmdd', StrToDateTime(TobTable.getString(champ)));
            if not IsDateVide(valchamp) then
              valeur2 := (FormatDateTime('yyyymmdd', StrToDateTime(valchamp)));
          end
    else  begin
            Valeur1 := TobTable.getString(champ);
            Valeur2 := valchamp;
          end;
  end;
  result := false;
  if operateur = 'C' then
  begin
    if copy(Valeur1, 1, length(Valeur2)) = Valeur2 then result := true else result := false;
  end else if operateur = 'L' then
  begin
    if pos(Valeur2, Valeur1) <> 0 then result := true else result := false;
  end else if operateur = '<>' then
  begin
    if Valeur1 <> Valeur2 then result := true else result := false;
  end else if operateur = '=' then
  begin
    if Valeur1 = Valeur2 then result := true else result := false;
  end else if operateur = 'I' then
  begin
    if pos(Valeur1, Valeur2) <> 0 then result := true else result := false;
  end else if operateur = 'E' then
  begin
    val := Valeur2;
    val1 := ReadToKenSt(val);
    val2 := ReadToKenSt(val);
    if (Valeur1 >= val1) and (Valeur1 <= val2)
      then result := true else result := false;
    exit;
  end else if operateur = '<' then
  begin
    if Valeur1 < Valeur2 then result := true else result := false;
  end else if operateur = '<=' then
  begin
    if Valeur1 <= Valeur2 then result := true else result := false;
  end else if operateur = 'D' then
  begin
    if copy(Valeur1, 1, length(Valeur2)) <> Valeur2 then result := true else result := false;
  end else if operateur = 'M' then
  begin
    if pos(Valeur2, Valeur1) = 0 then result := true else result := false;
  end else if operateur = 'J' then
  begin
    if pos(Valeur1, Valeur2) = 0 then result := true else result := false;
  end else if operateur = 'G' then
  begin
    val := Valeur2;
    val1 := ReadToKenSt(val);
    val2 := ReadToKenSt(val);
    if (Valeur1 < val1) or (Valeur1 > val2)
      then result := true else result := false;
    exit;
  end else if operateur = '>' then
  begin
    if Valeur1 > Valeur2 then result := true else result := false;
  end else if operateur = '>=' then
  begin
    if Valeur1 >= Valeur2 then result := true else result := false;
  end;
end;
{***********A.G.L.***********************************************
Auteur  ...... : M.N. GARNIER
Créé le ...... : 23/12/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

function Rt_NotifieAlerte(TobAlertes, TobTable: Tob): boolean;
var
  Memo: TRichEdit;
  Liste: HTStringList;
  ThePanel: TPanel;
  Q: TQuery;
  i: integer;
  StDestinataire: string;
  TobMail: tob;
begin
  result := true;

  if TobAlertes.GetBoolean('YAL_MEMO') then
  begin
    if TobAlertes.FieldExists(ALERT_SANSECRAN) and TobAlertes.GetBoolean(ALERT_SANSECRAN) then
    begin { Hors écran : Ramène le message d'alerte dans la tob alerte afin d'être traité côté métier plus tard }
      Result := (TobAlertes.GetString('YAL_MODEBLOCAGE') = SansBlocage)
             or ((TobAlertes.GetString('YAL_MODEBLOCAGE') = Interrogative) and (RecupCodeEvent(TobAlertes) <> CodeSuppression));
      if not Result then
        TobAlertes.AddChampSupValeur(ALERT_TobMsg, localRichToString(TobAlertes.GetString('YAL_BLOCNOTE')));
    end
    else { Notification écran }
      Result := StrToBool_(AGLLanceFiche('Y', 'YALERTES_MEMO', '', '', 'TOBALERTES=' + IntToStr(LongInt(TobAlertes))))
  end;

  if (TobAlertes.GetBoolean('YAL_EMAIL')) and result then
  begin
    ThePanel := TPanel.Create(nil);
    ThePanel.Visible := False;
    ThePanel.ParentWindow := GetDesktopWindow;
    Memo := TRichEdit.Create(ThePanel);
    Memo.Parent := ThePanel;
    Memo.width := 525;
    StringTorich(Memo, TobAlertes.GetValue('YAL_BLOCNOTE'));
    StDestinataire := '';
    TobMail := TOB.create('les mails', nil, -1);
    Q := OpenSql('Select * from YALERTESMAIL where YAM_ALERTE="' + TobAlertes.GetString('YAL_ALERTE') + '"', TRUE, -1, 'YALERTESMAIL');
    TobMail.LoadDetailDB('YALERTESMAIL', '', '', Q, False);
    ferme(Q);
    if TobMail.detail.count <> 0 then
    begin
      for i := 0 to TobMail.detail.count - 1 do
        if TobMail.detail[i].GetString('YAM_MAIL') <> '' then
          StDestinataire := StDestinataire + TobMail.detail[i].GetString('YAM_MAIL') + ';'
    end;
  { code de l'alerte, le type d'évènement, l'objet, le code de l'objet
  14047 Au niveau du mul des alertes :
1. Enlever le champ Type d'alerte qui n'a plus lieu d'être.
2. Modifier le type de saisie du champ Notification de l'alerte (bouton radio ou liste de choix).
 }
    if StDestinataire <> '' then
    begin
      Liste := HTStringList.Create;
      Liste.Text := Memo.text;
      Liste.Insert(0, '');
      Liste.Insert(0, RT_DescroAlerte(TobTable));
      PGIEnvoiMail('Alerte : ' + TobAlertes.GetString('YAL_ALERTE') + ' ' + TobAlertes.GetString('YAL_LIBELLE') +
        ', ' + RechDom('YEVENEMENTS', RecupCodeEvent(TobAlertes), false) + ' sur ' + RechDom('WPREFIXE', TobAlertes.GetString('YAL_PREFIXE'), false),
        StDestinataire, '', Liste, '', true, 1, '', '');
      Liste.free;
    end;
    if ThePanel <> nil then
      ThePanel.Free;
    FreeAndNil(TobMail);
  end;

  if TobAlertes.GetString('YAL_MODEBLOCAGE') = Blocage then
    result := false;
  Rt_LogAlerte(TobAlertes);
end;

function RT_DescroAlerte(Tobtable : tob) : string;
var FieldName : MyArrayField;
  i			        : Integer;
  Table		      : String;
  iTable, iChamp: Integer;
  Ok            : Boolean;
	Mcd : IMCDServiceCOM;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
	Result := '';
  FieldName := WMakeFieldArray(Tobtable.NomTable);
  { Récupération du nom de la table }
  Table := Tobtable.NomTable;
  if Table = '' then exit;
  { Récupération du N° de la table dans V_PGI }
  iTable := TableToNum ( Table );
  {$IFDEF EAGLCLIENT}
    if High(V_Pgi.DEChamps[iTable]) <= 0 then
      ChargeDeChamps(iTable, TableToPrefixe(Table));
  {$ENDIF EAGLCLIENT}
  { Forme le where en collant les nom de champ et leur valeur }
  i := Low(FieldName) - 1; Ok := True;
  while (i < High(FieldName)) and Ok do
  begin
    Inc(i);

	  iChamp := ChampToNum(FieldName[i]);
    if iChamp <> -1 then
    begin
        if Result = '' then
          Result := 'Fiche concernée : '+MCD.ChampToLibelle(FieldName[i])+' = '+ Tobtable.GetString(FieldName[i])
        else
          Result := Result + ', ' + MCD.ChampToLibelle(FieldName[i])+' = '+ Tobtable.GetString(FieldName[i]);
    end
    else
      Ok := False;

  end;
  if not Ok then Result := '';
end;


{***********A.G.L.***********************************************
Auteur  ...... : M.N. GARNIER
Créé le ...... : 23/12/2004
Modifié le ... : 23/12/2004
Description .. :
Mots clefs ... :
*****************************************************************}

constructor The_Alerte.create;
begin
  inherited create;
  Form := nil;
  Prefixe := '';
  Evenement := '';
  TobTable := nil;
  TobsInitiales := nil;
end;

{***********A.G.L.***********************************************
Auteur  ...... : M.N. GARNIER
Créé le ...... : 23/12/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

class function The_Alerte.ExecuteAlerte(const F: TForm; const FPrefixe,
  FEVent: string; const FTob, FTobsInit: TOB; const A_OM: TOM = nil): Boolean;
begin
  result := true;
  if assigned(FMonAlerte) then
  begin
    with FMonAlerte do
    begin
      Form := F;
      Prefixe := FPrefixe;
      Evenement := FEvent;
      TobTable := FTob;
      TobsInitiales := FTobsInit;
      FOM := A_OM;
      result := Rt_VerifierAlertes(FMonAlerte);
    end;
  end;
end;

function The_Alerte.GetOM: TOM;
begin
  if Assigned(FOM) then
    Result := FOM
  else if Assigned(Form) then
  begin
    Result := TFFiche(Form).OM;
    if (Form is TFSaisieList) then OM := TFSaisieList(Form).OM;
  end
  else
    Result := nil
end;


function AGLExecuteAlerteLoad(parms: array of variant; nb: integer): variant;
var F: TForm;
  OM: TOM;
begin
  result := false;
  F := TForm(Longint(Parms[0]));
  OM := nil;
  if (F is TFFiche) then OM := TFFiche(F).OM else
    if (F is TFFicheGrid) then OM := TFFicheGrid(F).OM else
      if (F is TFFicheListe) then OM := TFFicheListe(F).OM else
      if (F is TFSaisieList) then OM := TFSaisieList(F).OM;
  if not Assigned(OM) then exit;
  if AlerteActive(TableToPrefixe(OM.TableName)) then
    Result := ExecuteAlerteLoad(F, true);
end;

function AGLExecuteAlerteUpdate(parms: array of variant; nb: integer): variant;
var F: TForm;
  OM: TOM;
begin
  result := false;
  OM := nil;
  F := TForm(Longint(Parms[0]));
  if (F is TFFiche) then OM := TFFiche(F).OM else
    if (F is TFFicheGrid) then OM := TFFicheGrid(F).OM else
      if (F is TFFicheListe) then OM := TFFicheListe(F).OM else
      if (F is TFSaisieList) then OM := TFSaisieList(F).OM;
  if not Assigned(OM) then exit;
  if AlerteActive(TableToPrefixe(OM.TableName)) then
    Result := ExecuteAlerteUpdate(F, true);
end;

function AGLExecuteAlerteDelete(parms: array of variant; nb: integer): variant;
var F: TForm;
  OM: TOM;
begin
  result := false;
  OM := nil;
  F := TForm(Longint(Parms[0]));
  if (F is TFFiche) then OM := TFFiche(F).OM else
    if (F is TFFicheGrid) then OM := TFFicheGrid(F).OM else
      if (F is TFFicheListe) then OM := TFFicheListe(F).OM else
      if (F is TFSaisieList) then OM := TFSaisieList(F).OM;
  if not Assigned(OM) then exit;
  if AlerteActive(TableToPrefixe(OM.TableName)) then
    Result := ExecuteAlertedelete(F, true);
end;

procedure DoCreateTobFinale(F: TForm);
var OM: TOM;
  nomTable: string;
begin
  OM := nil;
  if (F is TFFiche) then OM := TFFiche(F).OM else
    if (F is TFFicheGrid) then OM := TFFicheGrid(F).OM else
      if (F is TFFicheListe) then OM := TFFicheListe(F).OM else
        if (F is TFSaisieList) then OM := TFSaisieList(F).OM;
  if not Assigned(OM) then exit;
  nomTable := OM.TableName;
  if (F is TFFiche) then
  begin
    if not Assigned(TFFiche(F).TobFinale) then
    begin
      TfFiche(F).TobFinale := TOB.create(nomTable, nil, -1);
      TfFiche(F).TobFinale.GetEcran(F, nil, true);
    end;
  end
  else
    if (F is TFFicheGrid) then
    begin
      if not Assigned(TFFicheGrid(F).TobFinale) then
      begin
        TFFicheGrid(F).TobFinale := TOB.create(nomTable, nil, -1);
        TFFicheGrid(F).TobFinale.GetEcran(F, nil, true);
      end;
    end
    else
      if (F is TFFicheListe) then
      begin
        if not Assigned(TFFicheListe(F).TobFinale) then
        begin
          TFFicheListe(F).TobFinale := TOB.create(nomTable, nil, -1);
          TFFicheListe(F).TobFinale.GetEcran(F, nil, true);
        end;
      end
    else
      if (F is TFSaisieList) then
      begin
        if not Assigned(TFSaisieList(F).TobFinale) then
        begin
          TFSaisieList(F).TobFinale := TOB.create(nomTable, nil, -1);
          TFSaisieList(F).TobFinale.GetEcran(F, nil, true);
        end;
      end;
end;

procedure DoCreateTobOrigine(F: TForm);
var OM: TOM;
  nomTable: string;
begin
  OM := nil;
  if (F is TFFiche) then OM := TFFiche(F).OM else
    if (F is TFFicheGrid) then OM := TFFicheGrid(F).OM else
      if (F is TFFicheListe) then OM := TFFicheListe(F).OM else
        if (F is TFSaisieList) then OM := TFSaisieList(F).OM;
  if not Assigned(OM) then exit;
  nomTable := OM.TableName;
  if (F is TFFiche) then
  begin
    if not Assigned(TFFiche(F).TobOrigine) then
    begin
      TfFiche(F).TobOrigine := TOB.create(nomTable, nil, -1);
      TfFiche(F).TobOrigine.GetEcran(F, nil, true);
    end;
  end
  else
    if (F is TFFicheGrid) then
    begin
      if not Assigned(TFFicheGrid(F).TobOrigine) then
      begin
        TFFicheGrid(F).TobOrigine := TOB.create(nomTable, nil, -1);
        TFFicheGrid(F).TobOrigine.GetEcran(F, nil, true);
      end;
    end
    else
      if (F is TFFicheListe) then
      begin
        if not Assigned(TFFicheListe(F).TobOrigine) then
        begin
          TFFicheListe(F).TobOrigine := TOB.create(nomTable, nil, -1);
          TFFicheListe(F).TobOrigine.GetEcran(F, nil, true);
        end;
      end
    else
      if (F is TFSaisieList) then
      begin
        if not Assigned(TFSaisieList(F).TobOrigine) then
        begin
          TFSaisieList(F).TobOrigine := TOB.create(nomTable, nil, -1);
          TFSaisieList(F).TobOrigine.GetEcran(F, nil, true);
        end;
      end;

end;

function ExecuteAlerteLoad(F: TForm; DoFree: boolean): boolean;
begin
  Result := true;
  if DoFree then
  begin
    DoFreeTobOrigine(F);
    DoFreeTobFinale(F);
  end;
  DoCreateTobFinale(F);
  DoCreateTobOrigine(F);

  if (F is TFFiche) then
    Result := FMonAlerte.ExecuteAlerte(F,
{$IFNDEF EAGLCLIENT}
    TFFiche(F).TableName,
{$ELSE}
    TableToPrefixe(TFFiche(F).TableName),
{$ENDIF EAGLCLIENT}
    CodeOuverture + ';' + CodeDateAnniv, TFFiche(F).TobFinale, TFFiche(F).TobOrigine)
  else
    if (F is TFFicheGrid) then
      Result := FMonAlerte.ExecuteAlerte(F,
{$IFNDEF EAGLCLIENT}
      TFFicheGrid(F).TableName,
{$ELSE}
      TableToPrefixe(TFFicheGrid(F).TableName),
{$ENDIF EAGLCLIENT}
      CodeOuverture + ';' + CodeDateAnniv, TFFicheGrid(F).TobFinale, TFFicheGrid(F).TobOrigine)
    else
      if (F is TFFicheListe) then
        Result := FMonAlerte.ExecuteAlerte(F,
{$IFNDEF EAGLCLIENT}
        TFFicheListe(F).TableName,
{$ELSE}
        TableToPrefixe(TFFicheListe(F).TableName),
{$ENDIF EAGLCLIENT}
        CodeOuverture + ';' + CodeDateAnniv, TFFicheListe(F).TobFinale, TFFicheListe(F).TobOrigine)
    else
      if (F is TFSaisieList) then
        Result := FMonAlerte.ExecuteAlerte(F, TableToPrefixe(TFSaisieList(F).TableName), CodeOuverture + ';' + CodeDateAnniv, TFSaisieList(F).TobFinale, TFSaisieList(F).TobOrigine)
end;

function ExecuteAlerteUpdate(F: TForm; DoFree: boolean): boolean;
var OM: TOM;
begin
  Result := true;
  if DoFree then
    DoFreeTobFinale(F);
  DoCreateTobFinale(F);
  OM := nil;
  if (F is TFFiche) then OM := TFFiche(F).OM else
    if (F is TFFicheGrid) then OM := TFFicheGrid(F).OM else
      if (F is TFFicheListe) then OM := TFFicheListe(F).OM else
      if (F is TFSaisieList) then OM := TFSaisieList(F).OM;
  if not Assigned(OM) then exit;

  if (F is TFFiche) then
    Result := FMonAlerte.ExecuteAlerte(F,
{$IFNDEF EAGLCLIENT}
    TFFiche(F).TableName,
{$ELSE EAGLCLIENT}
    TableToPrefixe(TFFiche(F).TableName),
{$ENDIF EAGLCLIENT}
    iif((GetStateFicheAlerte(F) = dsInsert), CodeCreation, CodeModification + ';' + CodeModifChamps)
      , TFFiche(F).TobFinale, TFFiche(F).TobOrigine)
  else
    if (F is TFFicheGrid) then
      Result := FMonAlerte.ExecuteAlerte(F,
{$IFNDEF EAGLCLIENT}
      TFFicheGrid(F).TableName,
{$ELSE EAGLCLIENT}
      TableToPrefixe(TFFicheGrid(F).TableName),
{$ENDIF EAGLCLIENT}
       iif((GetStateFicheAlerte(F) = dsInsert), CodeCreation, CodeModification + ';' + CodeModifChamps)
        , TFFicheGrid(F).TobFinale, TFFicheGrid(F).TobOrigine)
    else
      if (F is TFFicheListe) then
        Result := FMonAlerte.ExecuteAlerte(F,
{$IFNDEF EAGLCLIENT}
        TFFicheListe(F).TableName,
{$ELSE EAGLCLIENT}
        TableToPrefixe(TFFicheListe(F).TableName),
{$ENDIF EAGLCLIENT}
        iif((GetStateFicheAlerte(F) = dsInsert), CodeCreation, CodeModification + ';' + CodeModifChamps)
          , TFFicheListe(F).TobFinale, TFFicheListe(F).TobOrigine)
    else
      if (F is TFSaisieList) then
        Result := FMonAlerte.ExecuteAlerte(F, TableToPrefixe(TFSaisieList(F).TableName),
        iif((GetStateFicheAlerte(F) = dsInsert), CodeCreation, CodeModification + ';' + CodeModifChamps)
          , TFSaisieList(F).TobFinale, TFSaisieList(F).TobOrigine);
end;

function ExecuteAlerteDelete(F: TForm; DoFree: boolean): boolean;
begin
  Result := true;
  if DoFree then
    DoFreeTobFinale(F);
  DoCreateTobFinale(F);

  if (F is TFFiche) then
    Result := FMonAlerte.ExecuteAlerte(F,
{$IFNDEF EAGLCLIENT}
    TFFiche(F).TableName,
{$ELSE EAGLCLIENT}
    TableToPrefixe(TFFiche(F).TableName),
{$ENDIF EAGLCLIENT}
     CodeSuppression, TFFiche(F).TobFinale, TFFiche(F).TobOrigine)
  else
    if (F is TFFicheGrid) then
      Result := FMonAlerte.ExecuteAlerte(F,
{$IFNDEF EAGLCLIENT}
      TFFicheGrid(F).TableName,
{$ELSE EAGLCLIENT}
      TableToPrefixe(TFFicheGrid(F).TableName),
{$ENDIF EAGLCLIENT}
       CodeSuppression, TFFicheGrid(F).TobFinale, TFFicheGrid(F).TobOrigine)
    else
      if (F is TFFicheListe) then
        Result := FMonAlerte.ExecuteAlerte(F,
{$IFNDEF EAGLCLIENT}
        TFFicheListe(F).TableName,
{$ELSE EAGLCLIENT}
        TableToPrefixe(TFFicheListe(F).TableName),
{$ENDIF EAGLCLIENT}
         CodeSuppression, TFFicheListe(F).TobFinale, TFFicheListe(F).TobOrigine)
    else
      if (F is TFSaisieList) then
        Result := FMonAlerte.ExecuteAlerte(F, TableToPrefixe(TFSaisieList(F).TableName), CodeSuppression, TFSaisieList(F).TobFinale, TFSaisieList(F).TobOrigine)
end;

procedure DoFreeTobFinale(F: TForm);
begin
  if (F is TFFiche) and Assigned(TFFiche(F).TobFinale) then
  begin
    TFFiche(F).TobFinale.free;
    TFFiche(F).TobFinale := nil;
  end;
  if (F is TFFicheListe) and Assigned(TFFicheListe(F).TobFinale) then
  begin
    TFFicheListe(F).TobFinale.free;
    TFFicheListe(F).TobFinale := nil;
  end;

  if (F is TFFicheGrid) and Assigned(TFFicheGrid(F).TobFinale) then
  begin
    TFFicheGrid(F).TobFinale.free;
    TFFicheGrid(F).TobFinale := nil;
  end;
end;

procedure DoFreeTobOrigine(F: TForm);
begin
  if (F is TFFiche) and Assigned(TFFiche(F).TobOrigine) then
  begin
    TFFiche(F).TobOrigine.free;
    TFFiche(F).TobOrigine := nil;
  end;
  if (F is TFFicheListe) and Assigned(TFFicheListe(F).TobOrigine) then
  begin
    TFFicheListe(F).TobOrigine.free;
    TFFicheListe(F).TobOrigine := nil;
  end;

  if (F is TFFicheGrid) and Assigned(TFFicheGrid(F).TobOrigine) then
  begin
    TFFicheGrid(F).TobOrigine.free;
    TFFicheGrid(F).TobOrigine := nil;
  end;
end;

function ChargeTobAlerte(ObjAlerte: The_Alerte; bAlerteCroisee, bAlerte7: boolean): tob;
var i: integer;
  TobAlertes, TobAlertesF, TobFilleAlerte: tob;
  stEvent, ListeEvent: string;
  iNum: integer;
  ok: boolean;
begin
  TobAlertes := TOB.create('les alertes', nil, -1);
  for i := 0 to Pred(VH_EntPgi.TobAlertes.Detail.Count) do
  begin
    TobAlertesF := VH_EntPgi.TobAlertes.Detail[i];
    { test du ou des prefixes }
    if not bAlerteCroisee then
      if TobAlertesF.GetString('YAL_PREFIXE') <> ObjAlerte.Prefixe then
        continue;
    if bAlerteCroisee and not bAlerte7 then
      if (TobAlertesF.GetString('YAL_PREFIXE') <> ObjAlerte.Prefixe) and
        (TobAlertesF.GetString('YAL_PREFIXE') <> 'T') then
        continue;
    if bAlerteCroisee and not bAlerte7 then
      if (TobAlertesF.GetString('YAL_PREFIXE') <> ObjAlerte.Prefixe) and
        (TobAlertesF.GetString('YAL_PREFIXE') <> 'T') and
        (TobAlertesF.GetString('YAL_PREFIXE') <> 'GA') then
        continue;
    { test du ou des évènements }
    ok := false;
    ListeEvent := ObjAlerte.evenement;
    repeat
      { transfo du No }
      stEvent := ReadToKenSt(ListeEvent);
      if stEvent = '' then break;
      INum := Ord(stEvent[1]);
      if INum < 65 then
        INum := StrToInt(stEvent[1])
      else
        INum := Ord(stEvent[1]) - 55;
      stEvent := IntToStr(INum);
      if TobAlertesF.GetBoolean('YAL_EVENEMENT' + stEvent) then ok := true;
    until stEvent = '';

    if ok then
    begin
      TobFilleAlerte := TOB.Create('la fille alerte', TobAlertes, -1);
      TobFilleAlerte.Dupliquer(TobAlertesF, False, True);
    end;
  end;
  Result := TobAlertes;
end;

function AlerteActive(prefixe: string): boolean;
var TobAlerte: tob;
  ContPgi, stContexte, ContAlerte: string;
begin
  { les contextes paramétrés dans l'alerte doivent être présents dans les contextes de l'exe }
  result := false;
  VH_EntPgi.TobTableAlertes.Load;
  VH_EntPgi.TobAlertes.Load;
  VH_EntPgi.TobTablesLiees.Load;
  TobAlerte := VH_EntPgi.TobTableAlertes.FindFirst(['YTA_PREFIXE', 'YTA_ACTIF'],
    [prefixe, 'X'], false);
  if not Assigned(TobAlerte) then exit;
  result := true;
  if TobAlerte.GetString('YTA_CONTEXTES') = '<<Tous>>' then exit;
  ContPgi := AglContextesToStr(V_PGI.PGIContexte);
  ContAlerte := TobAlerte.GetString('YTA_CONTEXTES');
  repeat
    stContexte := ReadToKenSt(ContAlerte);
    if (stContexte <> '') and (pos(stContexte, ContPgi) = 0) then
      result := false;
  until (stContexte = '') or (result = false);
end;

function ExisteAlerteActive: boolean;
begin
  result := false;
  VH_EntPgi.TobTableAlertes.Load(true);
  VH_EntPgi.TobTablesLiees.Load(true);
  if Assigned(VH_EntPgi.TobTableAlertes.FindFirst(['YTA_ACTIF'], ['X'], false)) then result := true;
end;

function GetStateFicheAlerte(F: TForm): variant;
var
  DS: TDataSet;
begin
  if assigned(F) then
  begin
    DS := nil;
{$IFDEF EAGLCLIENT}
    if F is TFFiche then
      DS := TDataset(TFFiche(F).QFiche)
    else if F is TFFicheListe then
      DS := TDataset(TFFicheListe(F).Ta)
    else if F is TFSaisieList then
      DS := TFSaisieList(F).LeFiltre.DataSet;
{$ELSE}
    if F is TFFiche then
      DS := TDataset(F.FindComponent('QFiche'))
    else if F is TFFicheListe then
      DS := TDataset(F.FindComponent('Ta'))
    else if F is TFSaisieList then
      DS := TFSaisieList(F).LeFiltre.DataSet;
{$ENDIF}
    if assigned(DS) then
      result := DS.State
    else
      Result := dsInactive;
  end
  else
    Result := dsInactive;
end;

procedure Rt_LogAlerte(TobAlertDeclench: Tob);
var tobIndic: tob;
  Q: TQuery;
  maxI: integer;
  TobAlertes, TobAlertes2: tob;
begin
  maxI := 0;
  Q := OpenSQL('SELECT MAX(GIR_NUMERO) AS MAXNUM FROM ' + nomIndic, true);
  if not Q.EOF then
    maxI := Q.FindField('MAXNUM').AsInteger;
  Ferme(Q);
  if maxI = 0 then
    maxI := 1
  else
    maxI := maxI + 1;
  if (not TobAlertDeclench.GetBoolean('YAL_EVENEMENT6')) and
    (not TobAlertDeclench.GetBoolean('YAL_EVENEMENT7')) then
    TobAlertes := TobAlertDeclench
  else
  begin
    { on recherche les tob des alertes croisées }
    TobAlertes := VH_EntPgi.TobAlertes.FindFirst(['YAD_ALERTE', 'YAD_OBJET'],
      [TobAlertDeclench.GetString('YAL_ALERTE'), 'GP'], false);
    if not assigned(TobAlertes) then TobAlertes := TobAlertDeclench;
  end;
  tobIndic := TOB.create(nomIndic, nil, -1);
  tobIndic.initValeurs;
  tobIndic.SetInteger('GIR_NUMERO', maxI);
  tobIndic.SetString('GIR_INDICATEUR', TobAlertes.GetString('YAL_ALERTE'));
  tobIndic.SetDateTime('GIR_DATEEXECUT', Now);
  tobIndic.SetBoolean('GIR_ALERTE', True);
  if TobAlertes.GetString('YAU_CHAMP') <> '' then
    tobIndic.SetString('GIR_NOMCHAMP', TobAlertes.GetString('YAU_CHAMP'))
  else
  begin
    if TobAlertes.GetString('YAD_CHAMP1') <> '' then
    begin
      tobIndic.SetString('GIR_NOMCHAMP', TobAlertes.GetString('YAD_CHAMP1'));
      tobIndic.SetString('GIR_OPERATEUR', TobAlertes.GetString('YAD_OPER1'));
      tobIndic.SetString('GIR_VALCOND', TobAlertes.GetString('YAD_VAL1'));
    end;
    if TobAlertes.GetString('YAD_CHAMP2') <> '' then
    begin
      tobIndic.SetString('GIR_NOMCHAMP2', TobAlertes.GetString('YAD_CHAMP2'));
      tobIndic.SetString('GIR_OPERATEUR2', TobAlertes.GetString('YAD_OPER2'));
      tobIndic.SetString('GIR_VALCOND2', TobAlertes.GetString('YAD_VAL2'));
    end
    else
      { cas des alertes croisées : on ne renseigne que la première condition des 3 croisements possibles
        cad Pièces / Tiers / Articles, et pour ce on va les rechercher en mémoire }
      if TobAlertes.GetBoolean('YAL_EVENEMENT7') then
      begin
        TobAlertes2 := VH_EntPgi.TobAlertes.FindFirst(['YAD_ALERTE', 'YAD_OBJET'],
          [TobAlertDeclench.GetString('YAL_ALERTE'), 'T'], false);
        if assigned(TobAlertes2) then
        begin
          tobIndic.SetString('GIR_NOMCHAMP2', TobAlertes2.GetString('YAD_CHAMP1'));
          tobIndic.SetString('GIR_OPERATEUR2', TobAlertes2.GetString('YAD_OPER1'));
          tobIndic.SetString('GIR_VALCOND2', TobAlertes2.GetString('YAD_VAL1'));
        end
      end;
    if TobAlertes.GetString('YAD_CHAMP3') <> '' then
    begin
      tobIndic.SetString('GIR_NOMCHAMP3', TobAlertes.GetString('YAD_CHAMP3'));
      tobIndic.SetString('GIR_OPERATEUR3', TobAlertes.GetString('YAD_OPER3'));
      tobIndic.SetString('GIR_VALCOND3', TobAlertes.GetString('YAD_VAL3'));
    end
    else
      { cas des alertes croisées : on ne renseigne que la première condition des 3 croisements possibles
        cad Pièces / Tiers / Articles, et pour ce on va les rechercher en mémoire }
      if TobAlertes.GetBoolean('YAL_EVENEMENT7') then
      begin
        TobAlertes2 := VH_EntPgi.TobAlertes.FindFirst(['YAD_ALERTE', 'YAD_OBJET'],
          [TobAlertDeclench.GetString('YAL_ALERTE'), 'GA'], false);
        if assigned(TobAlertes2) then
        begin
          tobIndic.SetString('GIR_NOMCHAMP3', TobAlertes2.GetString('YAD_CHAMP1'));
          tobIndic.SetString('GIR_OPERATEUR3', TobAlertes2.GetString('YAD_OPER1'));
          tobIndic.SetString('GIR_VALCOND3', TobAlertes2.GetString('YAD_VAL1'));
        end
      end;
  end;
  tobIndic.SetString('GIR_EMAIL', TobAlertes.GetString('YAL_EMAIL'));
  tobIndic.SetString('GIR_MEMO', TobAlertes.GetString('YAL_MEMO'));
  tobIndic.SetString('GIR_MODEBLOCAGE', TobAlertes.GetString('YAL_MODEBLOCAGE'));
  tobIndic.SetString('GIR_PREFIXE', TobAlertes.GetString('YAL_PREFIXE'));
  tobIndic.SetString('GIR_VERSION', TobAlertes.GetString('YAL_VERSIONALERTE'));
  { il s'agit d'une alerte interractive }
  tobIndic.SetString('GIR_TYPEALERTE', 'I');

  tobIndic.SetString('GIR_EVENEMENTALERT', RecupCodeEvent(TobAlertes));
  tobIndic.InsertDB(nil);
  tobIndic.free;
end;

function RecupCodeEvent(TobAlertes: tob): string;
var i: integer;
begin
  Result := '';
  for i := 1 to nbEvents do
    if TobAlertes.GetBoolean('YAL_EVENEMENT' + IntToStr(i)) then
    begin
      Result := CodeEvent[i];
      break;
    end;
end;

initialization
  FMonAlerte := The_Alerte.Create();
{$IFNDEF ERADIO}
  RegisterAglFunc('ExecuteAlerteUpdate', TRUE, 0, AGLExecuteAlerteUpdate);
  RegisterAglFunc('ExecuteAlerteLoad', TRUE, 0, AGLExecuteAlerteLoad);
  RegisterAglFunc('ExecuteAlerteDelete', TRUE, 0, AGLExecuteAlerteDelete);
{$ENDIF}
finalization
  if assigned(FMonAlerte) then
    FreeAndNil(FMonAlerte);
end.
