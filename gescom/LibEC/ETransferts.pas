unit ETransferts;

interface
uses UTOB, ETransUtil;


//procedure TestBool;
procedure ExportAllTables;
procedure ImportAllTables;
procedure BalanceTradTablette;
procedure BalanceLeDispo;
procedure BalanceLesDevises;
procedure BalanceArticleTiers;
function ImporteECommande(Quiet : boolean = false; FromFiles : boolean = false; ToFiles : boolean = false; BouinageTiers : boolean = false; BouinageAddx : boolean = true) : boolean;
//function ImporteECommandeToFiles : boolean;

type TConvTable = Array of array[0..1] of Variant;
     TAddxConvAndGetAutoRegl = class
     private
       ConverTable, TiersConvTab : TConvTable;
       function GetNumeroConverti(Tab : TConvTable; Num : Variant) : Variant;

     public
       PiecesAutoRegl : TOB;
       GotToBouineZeTiers,
       GotToBouineZeAddress : boolean;
       function CheckDoublons(T : TOB) : integer;
       function ConvertAddxNbsAndGetAutoReglAndConvertTiers(T : TOB) : integer;
       function ConvertTiersLigne(T : TOB) : integer;
       function BouineLesTiers(T : TOB) : integer;

     end;

     TPiedPorc = class
     public
       function CreateMissingPorts(T : TOB) : integer;

     end;

const LastMSMQErrorMsg : String = '';


implementation
uses Controls, dbtables, HCtrls, HMsgBox, Ent1, FE_Main, CmdImport, sysutils, HEnt1, ParamSoc{,
     classes, ebizutil};

const {PGI2E : Array[1..25] of String =
              ('ADRESSES','ARTICLE', 'ARTICLELIE', 'CHANCELL', 'CHOIXCOD', 'COMMERCIAL',
               'COMMUN', 'CONTACT', 'DECOMBOS', 'DEVISE', 'DIMENSION', 'DIMMASQUE',
               'JUTYPECIVIL', 'LIENSOLE', 'MEA', 'MODEPAIE', 'MODEREGL', 'PARAMSOC',
               'PAYS', 'TARIF', 'TIERS', 'GTRADARTICLE', 'TRADDICO', 'TRADTABLETTE', 'TXCPTTVA');}
      PGI2E : Array[1..11] of String = // Version Light
              ('ARTICLE', 'ARTICLELIE', 'CHOIXCOD','DIMENSION', 'DIMMASQUE','LIENSOLE',
               'PARAMSOC', 'TARIF', 'GTRADARTICLE', 'TRADDICO', 'TRADTABLETTE');
      WhereExport : Array[0..0] of TTableWhere = ((Table:'ARTICLE'; Where:'GA_INVISIBLEWEB<>"X"'));
      E2PGI : Array[1..6] of String =
              ('ADRESSES', 'CONTACT', 'ELIGNE', 'EPIECE', 'ETIERS', 'RIB');
      ExpCaption = 'Transfert PGI -> E';
      ImpCaption = 'Transfert E -> PGI';

(*
procedure TestBool;
var S : TStringStream;
    T : TOB;
    Q : TQuery;
begin
   S := TStringStream.Create('');

   Q := OpenSQL('SELECT * FROM LIENSOLE WHERE LO_IDENTIFIANT="F0001" AND LO_EMPLOIBLOB="003"', true);
   T := TOB.Create('Mama_LIENSOLE', nil, -1);
   T.LoadDetailDB('LIENSOLE', '', '', Q, true);
   Ferme(Q);
   DeRTFizeLiensOLE(T);
   T.SaveToBinStream(S, false, true, false, false);
   T.Free;

   T := TOB.Create('LIENSOLE', nil, -1);
   S.Seek(0, soFromBeginning);
   BinStringToTOB(S.DataString, T);
   PGIBox(T.Detail[0].Detail[0].GetValue('LO_OBJET'), '');
   T.Free;

   S.Free;
end;*)

procedure ExportAllTables;
var Exp : TETransfert;
    b : boolean;
begin
  GetParams;
  if PGIAsk('Exporter toutes les tables vers e-business?', 'Transfert') = mrNo then exit;

  Exp := TETransfert.Create;
  b := Exp.ExportFullTables(PGI2E, false, DefaultFormat, WhereExport);
  Exp.Free;

  if b then PGIInfo('Export effectué avec succès', 'Transfert')
       else PGIInfo('Export abandonné', 'Transfert');
end;

procedure ImportAllTables;
var Exp : TETransfert;
    b : boolean;
begin
  GetParams;
  if PGIAsk('Importer toutes les tables depuis e-business?', 'Transfert') = mrNo then exit;

  Exp := TETransfert.Create;
  b := Exp.ImportFullTables(E2PGI, false, DefaultFormat);
  Exp.Free;

  if b then PGIInfo('Import effectué avec succès', 'Transfert')
       else PGIInfo('Import abandonné', 'Transfert');
end;


procedure BalanceTradTablette;
var Exp : TETransfert;
begin
  GetParams;
  if PGIAsk('Exporter les traductions de tablettes vers e-business?', 'Transfert') = mrNo then exit;

  Exp := TETransfert.Create;
  Exp.ExportFullTables(['TRADTABLETTE'], true, DefaultFormat, []);
  if Exp.PostMQ then PGIInfo('Export effectué avec succès', 'Transfert')
                else PGIInfo('Une erreur est survenue pendant l''export :'#13+Exp.LastErrorMsg, 'Transfert');
  Exp.Free;
end;

procedure BalanceLeDispo;
var Exp : TETransfert;
const WE : TTableWhere = (Table:'DISPO'; Where:'GQ_CLOTURE="-"'); // DBR pour oracle, clé pas utilisée si test GQ_CLOTURE<>"X"
begin
  GetParams;
  if PGIAsk('Exporter le disponible des stocks vers e-business?', 'Transfert') = mrNo then exit;

  Exp := TETransfert.Create;
  Exp.ExportFullTables(['DISPO'], true, DefaultFormat, [WE]);
  if Exp.PostMQ then PGIInfo('Export effectué avec succès', 'Transfert')
                else PGIInfo('Une erreur est survenue pendant l''export :'#13+Exp.LastErrorMsg, 'Transfert');
  Exp.Free;
end;

procedure BalanceLesDevises;
var Exp : TETransfert;
    WE : TTableWhere;
begin
  GetParams;
  if PGIAsk('Exporter les devises vers e-business?', 'Transfert') = mrNo then exit;

  Exp := TETransfert.Create;
  WE.Table := 'CHANCELL'; WE.Where := 'H_DATECOURS>="'+USDateTime(VH^.EnCours.Deb)+'"';
  Exp.ExportFullTables(['CHANCELL'], true, DefaultFormat, [WE]);
  if Exp.PostMQ then PGIInfo('Export effectué avec succès', 'Transfert')
                else PGIInfo('Une erreur est survenue pendant l''export :'#13+Exp.LastErrorMsg, 'Transfert');
  Exp.Free;
end;

procedure BalanceArticleTiers;
var Exp : TETransfert;
begin
  GetParams;
  if PGIAsk('Exporter les liens article-tiers vers e-business?', 'Transfert') = mrNo then exit;

  Exp := TETransfert.Create;
  Exp.ExportFullTables(['ARTICLETIERS'], true, DefaultFormat, []);
  if Exp.PostMQ then PGIInfo('Export effectué avec succès', 'Transfert')
                else PGIInfo('Une erreur est survenue pendant l''export :'#13+Exp.LastErrorMsg, 'Transfert');
  Exp.Free;
end;


function TAddxConvAndGetAutoRegl.GetNumeroConverti(Tab : TConvTable; Num : Variant) : Variant;
var i : integer;
begin
  result := Num;
  for i := low(Tab) to high(Tab) do
   if Tab[i, 0] = Num then
   begin result := Tab[i, 1];
         break; end;
end;

function TAddxConvAndGetAutoRegl.CheckDoublons(T : TOB) : integer;
var i, nad : integer;
    Q : TQuery;
begin
  result := 0;
  if not GotToBouineZeAddress then exit;

  // Anti écrasement d'adresses
  Q := OpenSQL('SELECT MAX(ADR_NUMEROADRESSE) FROM ADRESSES', true);
  if Q.EOF then nad := 0
           else nad := Q.Fields[0].AsInteger+1;
  Ferme(Q);

  for i := 0 to T.Detail.Count-1 do
  begin
    SetLength(ConverTable, Length(ConverTable)+1);
    ConverTable[High(ConverTable), 0] := T.Detail[i].GetValue('ADR_NUMEROADRESSE');
    ConverTable[High(ConverTable), 1] := nad+i;
    T.Detail[i].PutValue('ADR_NUMEROADRESSE', nad+i);
  end;
end;

function TAddxConvAndGetAutoRegl.ConvertAddxNbsAndGetAutoReglAndConvertTiers(T : TOB) : integer;
var i : integer;
begin
  result := 0;

  for i := 0 to T.Detail.Count-1 do
  begin
    if GotToBouineZeAddress then
    begin
       T.Detail[i].PutValue('EP_NUMADRESSELIVR', GetNumeroConverti(ConverTable, T.Detail[i].GetValue('EP_NUMADRESSELIVR')));
       T.Detail[i].PutValue('EP_NUMADRESSEFACT', GetNumeroConverti(ConverTable, T.Detail[i].GetValue('EP_NUMADRESSEFACT')));
    end;

    if ModeReglAuto(T.Detail[i].GetValue('EP_MODEREGLE'))
       then TOB.Create('EPIECE', PiecesAutoRegl, -1).Dupliquer(T.Detail[i], true, true);

    // Rajout : Bouinage Tiers
    if GotToBouineZeTiers then
    begin
       T.Detail[i].PutValue('EP_TIERS', GetNumeroConverti(TiersConvTab, T.Detail[i].GetValue('EP_TIERS')));
       T.Detail[i].PutValue('EP_TIERSLIVRE', GetNumeroConverti(TiersConvTab, T.Detail[i].GetValue('EP_TIERSLIVRE')));
       T.Detail[i].PutValue('EP_TIERSFACTURE', GetNumeroConverti(TiersConvTab, T.Detail[i].GetValue('EP_TIERSFACTURE')));
    end;
  end;
end;

function TAddxConvAndGetAutoRegl.ConvertTiersLigne(T : TOB) : integer;
var i : integer;
begin
  result := 0;
  if not GotToBouineZeTiers then exit;

  for i := 0 to T.Detail.Count-1 do
  begin
    T.Detail[i].PutValue('EL_TIERS', GetNumeroConverti(TiersConvTab, T.Detail[i].GetValue('EL_TIERS')));
    T.Detail[i].PutValue('EL_TIERSLIVRE', GetNumeroConverti(TiersConvTab, T.Detail[i].GetValue('EL_TIERSLIVRE')));
    T.Detail[i].PutValue('EL_TIERSFACTURE', GetNumeroConverti(TiersConvTab, T.Detail[i].GetValue('EL_TIERSFACTURE')));
  end;
end;

function TAddxConvAndGetAutoRegl.BouineLesTiers(T : TOB) : integer;
var i, nouvonumeral : integer;
begin
  result := 0;
  if not GotToBouineZeTiers then
  begin LogDebugMsg('Pas besoin de bouiner les tiers', true);
        exit; end;

  for i := 0 to T.Detail.Count-1 do
  begin
    if ExisteSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE="'+T.Detail[i].GetValue('ETI_AUXILIAIRE')+'"') then
    begin LogDebugMsg('(Bouinage Tiers) Le tiers existe déjà ('+T.Detail[i].GetValue('ETI_AUXILIAIRE')+')', true);
          Continue; end;

    T.Detail[i].PutValue('ETI_EAN', T.Detail[i].GetValue('ETI_TIERS'));
    nouvonumeral := GetParamSoc('SO_GCCOMPTEURTIERS')+1;
    SetParamSoc('SO_GCCOMPTEURTIERS', nouvonumeral);

    SetLength(TiersConvTab, Length(TiersConvTab)+1);
    TiersConvTab[High(TiersConvTab), 0] := T.Detail[i].GetValue('ETI_TIERS');
    TiersConvTab[High(TiersConvTab), 1] := IntToStr(nouvonumeral);
    LogDebugMsg('Tiers Bouiné ! '+TiersConvTab[High(TiersConvTab), 0]+' -> '+TiersConvTab[High(TiersConvTab), 1], true);

    T.Detail[i].PutValue('ETI_TIERS', IntToStr(nouvonumeral));
    T.Detail[i].PutValue('ETI_AUXILIAIRE', IntToStr(nouvonumeral)+'00');
  end;
end;


function TPiedPorc.CreateMissingPorts(T : TOB) : integer;
var i : integer;
    TypePort : String;
begin
  result := 0;
  for i := 0 to T.Detail.Count-1 do
  begin
    if not ExisteSQL('SELECT GPO_CODEPORT FROM PORT WHERE GPO_CODEPORT="'+T.Detail[i].GetValue('GPT_CODEPORT')+'"') then
    begin
      TypePort := T.Detail[i].GetValue('GPT_TYPEPORT');
      if TypePort = '' then TypePort := 'MT';

      with TOB.Create('PORT', nil, -1) do
      begin
        PutValue('GPO_CODEPORT', T.Detail[i].GetValue('GPT_CODEPORT'));
        PutValue('GPO_LIBELLE', T.Detail[i].GetValue('GPT_LIBELLE'));
        PutValue('GPO_TYPEPORT', TypePort);
        PutValue('GPO_PVHT', T.Detail[i].GetValue('GPT_TOTALHT'));
        PutValue('GPO_PVTTC', T.Detail[i].GetValue('GPT_TOTALTTC'));
        PutValue('GPO_COMPTAARTICLE', T.Detail[i].GetValue('GPT_COMPTAARTICLE'));
        PutValue('GPO_FAMILLETAXE1', T.Detail[i].GetValue('GPT_FAMILLETAXE1'));
        PutValue('GPO_FAMILLETAXE2', T.Detail[i].GetValue('GPT_FAMILLETAXE2'));
        PutValue('GPO_FAMILLETAXE3', T.Detail[i].GetValue('GPT_FAMILLETAXE3'));
        PutValue('GPO_FAMILLETAXE4', T.Detail[i].GetValue('GPT_FAMILLETAXE4'));
        PutValue('GPO_FAMILLETAXE5', T.Detail[i].GetValue('GPT_FAMILLETAXE5'));
        InsertDB(nil);
        Free;
      end;
    end;
  end;
end;


function ImporteECommande(Quiet : boolean = false; FromFiles : boolean = false; ToFiles : boolean = false; BouinageTiers : boolean = false; BouinageAddx : boolean = true) : boolean;
var Exp : TETransfert;
    TAC : TAddxConvAndGetAutoRegl;
    TPP : TPiedPorc;
    i : integer;
    ok : boolean;
    EPN : String;
    MereAdoptive : TOB;
begin
  if not Quiet then GetParams;

  Exp := TETransfert.Create;
  Exp.Quiet := Quiet;
  TAC := TAddxConvAndGetAutoRegl.Create;
  TAC.PiecesAutoRegl := TOB.Create('Mama_EPIECE', nil, -1);
  TAC.GotToBouineZeTiers := BouinageTiers;
  TAC.GotToBouineZeAddress := BouinageAddx;
  TPP := TPiedPorc.Create;
  if not FromFiles then LogDebugMsg('Récupération des données via MSMQ...', true);
  if not Quiet then Exp.CreateTT('Import de commandes', 'Récupération des données en cours...', 9, false, true);
  LogDebugMsg('(Bouinage Tiers activé)', true);

  result := false;
  try
    {$B-}
    if FromFiles or Exp.GetMQ then
    begin
      if ToFiles then LogDebugMsg('Sauvegarde des données dans des fichiers...', true)
                 else LogDebugMsg('Insertion des données dans les tables...', true);
      LastMSMQErrorMsg := '';
      Exp.SetTTValue(1);
      while Exp.FormatToTable('Import_ETIERS', true, not FromFiles, DefaultFormat, TAC.BouineLesTiers, ToFiles, true) do result := true;
      Exp.SetTTValue(2);
      while Exp.FormatToTable('Import_ADRESSES', false, not FromFiles, DefaultFormat, TAC.CheckDoublons, ToFiles, true) do result := true; // D'abord les adresses avec detection de n°s doublons et création d'une table de conversion
      Exp.SetTTValue(3);
      while Exp.FormatToTable('Import_PIECEADRESSE', false, not FromFiles, DefaultFormat, nil, ToFiles, true) do result := true;
      Exp.SetTTValue(4);
      while Exp.FormatToTable('Import_EPIECE', false, not FromFiles, DefaultFormat, TAC.ConvertAddxNbsAndGetAutoReglAndConvertTiers, ToFiles, true) do result := true; // Utiliser la tab2conversion pour bouiner les n°s addx
      Exp.SetTTValue(5);
      while Exp.FormatToTable('Import_ELIGNE', false, not FromFiles, DefaultFormat, TAC.ConvertTiersLigne, ToFiles, true) do result := true;
      Exp.SetTTValue(6);
      while Exp.FormatToTable('Import_PIEDPORT', false, not FromFiles, DefaultFormat, TPP.CreateMissingPorts, ToFiles, true) do result := true;
      Exp.SetTTValue(7);
      while Exp.FormatToTable('Import_PIEDBASE', false, not FromFiles, DefaultFormat, nil, ToFiles, true) do result := true;
      Exp.SetTTValue(8);

      if (not ToFiles) and (TAC.PiecesAutoRegl.Detail.Count > 0) then
      begin
         LogDebugMsg('Intégration automatique des commandes ('+inttostr(TAC.PiecesAutoRegl.Detail.Count)+' pièces)...', true);
         Exp.SetTTSubText('Intégration automatique des commandes...');

         //IntegreCommandes(TAC.PiecesAutoRegl, false);

         ok := true;
         MereAdoptive := TOB.Create('', nil, -1);
         for i := 0 to TAC.PiecesAutoRegl.Detail.Count-1 do  // Intégration une par une
         begin
            TAC.PiecesAutoRegl.Detail[i].ChangeParent(MereAdoptive, -1);
            EPN := MereAdoptive.Detail[0].GetValue('EP_NUMERO');
            try
              V_PGI.IoError := oeOK;
              IntegreCommandes(MereAdoptive, false, Quiet);
              if V_PGI.IoError <> oeOK then begin LogDebugMsg('[Intégration Auto |'+EPN+'] '+ioErrMsg(V_PGI.IoError, 'intégr'), false); ok := false; end;
            except
              on E : Exception do begin LogDebugMsg('[Intégration Auto |'+EPN+'] '+E.Message, false); ok := false; end;
            end;
            MereAdoptive.Detail[0].ChangeParent(TAC.PiecesAutoRegl, i);
         end;

         MereAdoptive.Free;
         if ok then LogDebugMsg('[Intégration Auto] '+ioErrMsg(V_PGI.IoError, 'intégr'), false);
      end;
      Exp.SetTTValue(9);
    end
     else LastMSMQErrorMsg := 'GetMQ Failed ('+Exp.LastErrorMsg+')';

  finally
    TAC.PiecesAutoRegl.Free;
    TAC.Free;
    TPP.Free;
    Exp.FreeTT;
    Exp.Free;

  end;

  if result and (not Quiet) then AGLLanceFiche('E','EIMPORTCMD','','','');
end;


end.
