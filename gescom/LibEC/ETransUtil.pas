unit ETransUtil;

interface
uses Forms, UTOB, ed_tools{$IFNDEF EAGLCLIENT}, dbtables, EasyMSMQ{$ENDIF};

type TTOBFormat = (tfTob, tfBin, tfXML);
     TTableWhere = record
        Table : String;
        Where : String;
      end;
{$IFNDEF EAGLCLIENT}
     TETransfert = class
     private
       FTP : TUseTOB;
       FTTAssigned : boolean;
       FTTSubText : string;
       FTTCanceled : boolean;

     public
       FMQin, FMQout : TMSMQueue;
       Quiet : boolean;
       LastErrorMsg : String;

       constructor Create;
       destructor Destroy; override;

       procedure CreateMQin;
       procedure CreateMQout;

       procedure TOBToFormat(var T : TOB; FFileName : String; ParMSMQ : boolean; Fmt : TTOBFormat);
       procedure QueryToFormat(Q : TQuery; FFileName : String; ParMSMQ : boolean; Fmt : TTOBFormat; TraitementSpecif : TUseTOB = nil);
       function TOBToTable(T : TOB; Convert : boolean) : integer;
       function FormatToTable(FFileName : String; Convert : boolean; ParMSMQ : boolean; Fmt : TTOBFormat; TraitementPrealable : TUseTOB = nil; SaveToFile : boolean = false; DeleteFiles : boolean = false) : boolean;

       function GetMQ : boolean;
       function PostMQ : boolean;

       function ExportFullTables(TableList : Array of String; ParMSMQ : boolean; Fmt : TTOBFormat; WhereList : Array of TTableWhere) : boolean;
       function ImportFullTables(TableList : Array of String; ParMSMQ : boolean; Fmt : TTOBFormat) : boolean;

       procedure CreateTT(Caption, Titre : string; MaxValue : integer; EnableCancel, DisabledMenu : boolean);
       procedure SetTTMaxValue(Max : integer);
       procedure SetTTValue(Val : integer);
       procedure SetTTSubText(St : String);
       function TTCanceled : boolean;
       procedure FreeTT;

     end;
{$ENDIF}
var LogDebugMsgProc : procedure(Msg : String; IsDebugMsg : boolean) of object;
procedure LogDebugMsg(Msg : String; IsDebugMsg : boolean);

function StrInArray(Str : String; const Arr : Array of String) : boolean;
procedure GetParams;
//function GetModeRegl(MR : String) : String;
function IsModeReglChq(GPMR : String) : boolean;
function ModeReglAuto(MR : String) : boolean;
procedure ETableToTable(SrcTOB, DestTOB : TOB);
procedure TableToETable(SrcTOB, DestTOB : TOB);
function DeRTFize(S : String; ParentForm : TForm = nil) : String;
procedure DeRTFizeLiensOLE(T : TOB);
procedure CloseDeRTFizator;
procedure DemoulineTarif;

const ExpFolder : String = 'Generation\';
      ImpFolder : String = 'Import\';
      InQueuePath : String = 'webpgi';
      OutQueuePath : String = 'pgiweb';
      GotToZip : boolean = false;
      Encrypted : boolean = false;
      LiensOLEQuandTiers : boolean = false;
      DoitNbSpizer : boolean = false;
      ValidationAuto : boolean = false;
      DefaultFormat : TTOBFormat = tfBin;

      ExpCaption : string = 'Transfert PGI -> E';
      ImpCaption : string = 'Transfert E -> PGI';

implementation
uses Classes, comctrls, SysUtils, HCtrls, HMsgBox, HEnt1, Ent1, EntGC, EBizUtil, ParamSoc,
     yTarifsCommun, TarifUtil;

const ConvertibleTables : Array[1..6] of string = ('ARTICLE', 'TIERS', 'EARTICLE', 'ELIGNE', 'EPIECE', 'ETIERS');

procedure LogDebugMsg(Msg : String; IsDebugMsg : boolean);
begin
   if Assigned(LogDebugMsgProc) then LogDebugMsgProc(Msg, IsDebugMsg);
end;

function StrInArray(Str : String; const Arr : Array of String) : boolean;
var i : integer;
begin
  result := false;
  for i := Low(Arr) to High(Arr) do
   if Arr[i] = Str then
   begin
     result := true;
     break;
   end;
end;

function TableAConvertir(Name : String) : boolean;
begin
  result := StrInArray(Name, ConvertibleTables);
end;

{$IFNDEF EAGLCLIENT}
constructor TETransfert.Create;
begin
  inherited;
{  FMQin := TMSMQueue.Create(InQueuePath);
  FMQin.Exclusif := true;
  FMQout := TMSMQueue.Create(OutQueuePath);
  FMQout.Encrypted := Encrypted; }
  Quiet := false;
  FTTAssigned := false;
end;

destructor TETransfert.Destroy;
begin
  if Assigned(FMQin) then FMQin.Free;
  if Assigned(FMQout) then FMQout.Free;
  inherited;
end;

procedure TETransfert.CreateMQin;
begin
  if Assigned(FMQin) then exit;
  FMQin := TMSMQueue.Create(InQueuePath);
  FMQin.Exclusif := true;
end;

procedure TETransfert.CreateMQout;
begin
  if Assigned(FMQout) then exit;
  FMQout := TMSMQueue.Create(OutQueuePath);
  FMQout.Encrypted := Encrypted;
end;

// Envoi d'une tob par fichier ou par MSMQ, en format Fmt, compressé ou non,
//  en convertissant au passage et si y'a besoin la table en E-table
procedure TETransfert.TOBToFormat(var T : TOB; FFileName : String; ParMSMQ : boolean; Fmt : TTOBFormat);
var Pfx, Ext, TN, StrXML : String;
    p : integer;
    TConv : TOB;
    FOut : TFileStream;
    E : boolean;
begin
  if T.Detail.Count = 0 then exit;

  TN := T.Detail[0].NomTable;
  if TableAConvertir(TN) then
  begin
    E := (TN[1] = 'E');
    if not E then
    begin
      SetTTSubText('Conversion vers table E-Commerce...');
      LogDebugMsg('Conversion vers table E-Commerce...', true);
      TN := 'E'+TN;
      Insert('E', FFileName, Pos('_', FFileName)+1);
    end else
    begin
      SetTTSubText('Conversion vers table PGI...');
      LogDebugMsg('Conversion vers table PGI...', true);
      Delete(TN,1,1);
      Delete(FFileName, Pos('_', FFileName)+1, 1);
    end;
    TConv := T;
    T := TOB.Create('Mama_'+TN, nil, -1);
    if not E then TableToETable(TConv, T)
             else ETableToTable(TConv, T);
    TConv.Free;
  end;

  case Fmt of
    tfTob : Ext := 'TOB';
    tfBin : Ext := 'binaire';
    tfXML : Ext := 'XML';
  end;
  if GotToZip then Ext := Ext + ' compressé';
  SetTTSubText('Export au format '+Ext+'...');
  LogDebugMsg('Export au format '+Ext+'...', true);

  case Fmt of
    tfTob : begin StrXML := TOBtoString(T, GotToZip); Ext := '.tob'; end;
    tfBin : begin StrXML := TOBtoBINString(T, GotToZip); Ext := '.bin'; end;
    tfXML : begin StrXML := TOBtoXMLString(T, GotToZip); Ext := '.xml'; end;
  end;
  if GotToZip then Ext := Ext + '.hz';

  if ParMSMQ then
  begin
    CreateMQout;
    LogDebugMsg('Ajout à la file MSMQ', true);
    FMQout.EnqueueMsg(FFileName+Ext, StrXML);
  end else
  begin
    Pfx := ''; p := 0;
    while FileExists(ExpFolder+Pfx+FFileName+Ext) do begin Pfx := inttostr(p)+'.'; inc(p); end;
    LogDebugMsg('Ecriture de '+ExpFolder+Pfx+FFileName+Ext, true);
    FOut := TFileStream.Create(ExpFolder+Pfx+FFileName+Ext, fmCreate);
    FOut.Write(StrXML[1], SizeOf(StrXML[1])*Length(StrXML));
    FOut.Free;
  end;
end;

// Envoi des résultats d'un query par fichier ou par MSMQ, en format Fmt, compressé ou non,
//  en convertissant au passage et si y'a besoin la table en E-table
procedure TETransfert.QueryToFormat(Q : TQuery; FFileName : String; ParMSMQ : boolean; Fmt : TTOBFormat; TraitementSpecif : TUseTOB = nil);
var TableName : String;
    T : TOB;
begin
  TableName := GetTableNameFromDataSet(Q);
  T := TOB.Create('Mama_'+TableName, nil, -1);

  SetTTSubText('Lecture de la table "'+TableToLibelle(TableName)+'"');
  T.LoadDetailDB(TableName, '', '', Q, false, true);

  if TTCanceled then exit;

  TOBToFormat(T, FFileName, ParMSMQ, Fmt);

  if Assigned(TraitementSpecif) then TraitementSpecif(T);

  T.Free;
end;


// Insère une tob dans sa table en convertissant au passage et si y'a besoin la E-table en table
function TETransfert.TOBToTable(T : TOB; Convert : boolean) : integer;
var TN : String;
    TConv : TOB;
    E : boolean;
begin
  result := 0;
  if T.Detail.Count = 0 then begin T.Free; exit; end;

  if Assigned(FTP) then FTP(T);

  TN := T.Detail[0].NomTable;
  if Convert and TableAConvertir(TN) then
  begin
    E := (TN[1] = 'E');
    if not E then
    begin
      SetTTSubText('Conversion vers table E-Commerce...');
      TN := 'E'+TN;
    end else
    begin
      SetTTSubText('Conversion vers table PGI...');
      Delete(TN,1,1);
    end;
    TConv := T;
    T := TOB.Create('Mama_'+TN, nil, -1);
    if not E then TableToETable(TConv, T)
             else ETableToTable(TConv, T);
    TConv.Free;
{    SetTTSubText('Conversion vers table PGI...');
    TConv := T;
    Delete(TN,1,1);
    T := TOB.Create('Mama_'+TN, nil, -1);
    ETableToTable(TConv, T);
    TConv.Free; }
  end;

  SetTTSubText('Mise à jour de la table "'+TableToLibelle(T.Detail[0].NomTable)+'"');

  T.InsertOrUpdateDBTable(true);
  T.Free;

  if TTCanceled then result := -1;
end;

// Insère un fichier/message MSMQ au format Fmt (compressé ou non) dans sa table
//  en convertissant au passage et si y'a besoin la E-table en table
function TETransfert.FormatToTable(FFileName : String; Convert : boolean; ParMSMQ : boolean; Fmt : TTOBFormat; TraitementPrealable : TUseTOB = nil; SaveToFile : boolean = false; DeleteFiles : boolean = false) : boolean;
var Pfx, Ext, StrXML : String;
    p : integer;
    SR : TSearchRec;
    FIn, FOut : TFileStream;
    FMsg : TMsg;
    T : TOB;
begin
  result := false;
  case Fmt of
    tfTob : Ext := 'TOB';
    tfBin : Ext := 'binaire';
    tfXML : Ext := 'XML';
  end;
  if GotToZip then Ext := Ext + ' compressé';
  SetTTSubText('Chargement des données "'+TableToLibelle(Copy(FFileName, Pos('_', FFileName)+1, Length(FFileName)))+'" ('+Ext+')...');
  FTP := TraitementPrealable;

  case Fmt of
    tfTob : Ext := '.tob';
    tfBin : Ext := '.bin';
    tfXML : Ext := '.xml';
  end;
  if GotToZip then Ext := Ext+'.hz';
  if ParMSMQ then
  begin
    CreateMQin;
    FMsg := FMQin.MsgQueue.FindMsg(FFileName+Ext);
    //if FMsg = nil then raise Exception.Create('Impossible de trouver le message '+FFileName+'.xml'+Ext);
    if FMsg = nil then exit;
    StrXML := FMsg.Body;
    FMQin.MsgQueue.Suppr(FMQin.MsgQueue.IndexOf(FMsg));
  end else
  begin
    // Ici bouinage inverse pour retrouver un fichier unicizé
    // Pas besoin de faire de while pour lire tous les fichiers ici puisque on appelle la methode avec un while
    Pfx := ImpFolder+FFileName+Ext;
    LogDebugMsg('Recherche 1 : "'+Pfx+'"', true);
    if not FileExists(Pfx) then
    begin
        if FindFirst(ImpFolder+'*.'+FFileName+Ext, 0, SR) = 0 then Pfx := ImpFolder+SR.Name
                                                              else Pfx := '';
        FindClose(SR);
    end;
    LogDebugMsg('Recherche 2 : "'+Pfx+'"', true);
    if Pfx = '' then exit;
    LogDebugMsg('Lecture de '+Pfx, true);
    FIn := TFileStream.Create(Pfx {ImpFolder+FFileName+Ext}, fmOpenRead);
    SetLength(StrXML, FIn.Size);
    FIn.Read(StrXML[1], FIn.Size);
    FIn.Free;
    if DeleteFiles then if not DeleteFile(Pfx) then LogDebugMsg('Impossible de supprimer '+Pfx, false);
  end;

  if SaveToFile then
  begin
    // test d'existence et rajoutage de prefixe pour unicité
    Pfx := ''; p := 0;
    while FileExists(ExpFolder+Pfx+FFileName+Ext) do begin Pfx := inttostr(p)+'.'; inc(p); end;
    LogDebugMsg('Ecriture de '+ExpFolder+Pfx+FFileName+Ext, true);
    FOut := TFileStream.Create(ExpFolder+Pfx+FFileName+Ext, fmCreate);
    FOut.Write(StrXML[1], SizeOf(StrXML[1])*Length(StrXML));
    FOut.Free;
  end else
  begin
      T := TOB.Create('moma', nil, -1);
      case Fmt of
        tfTob : raise Exception.Create('Impossible d''importer des données au format TOB');
        tfBin : BINStringToTOB(StrXML, T, GotToZip);
        tfXML : XMLStringToTOB(StrXML, T, GotToZip);
      end;
      if T.Detail.Count = 0 then exit;
      TOBToTable(T.Detail[0], Convert);
  end;
  result := true;
end;

// Récupere les messages de la file d'entrée
function TETransfert.GetMQ : boolean;
begin
  SetTTSubText('Réception des données par MSMQ...');
  CreateMQin;
  result := FMQin.ReceiveQueue;
  SetTTSubText('Terminé');
  LastErrorMsg := FMQin.ErrorMsg;
end;

// Envoie les messages de la file de sortie
function TETransfert.PostMQ : boolean;
begin
//  CreateMQout;
  result := true;  // Renvoie true si envoyé ou si rien à envoyer (false seulement si erreur d'envoi)
  if not Assigned(FMQout) then exit;
  if FMQout.MsgQueue.Count = 0 then exit;
  SetTTSubText('Envoi des données par MSMQ...');
  result := FMQout.SendQueue;
  SetTTSubText('Terminé');
  if result then FMQout.ClearQueue;
  LastErrorMsg := FMQout.ErrorMsg;
end;


// Exporte une ou plusieurs tables en fichiers ou en messages MSMQ
function TETransfert.ExportFullTables(TableList : Array of String; ParMSMQ : boolean; Fmt : TTOBFormat; WhereList : Array of TTableWhere) : boolean;
var i : integer;
    StWhere : String;
    Q : TQuery;

  function GetWhere(TableName : String) : String;
  var j : integer;
  begin
    result := '';
    for j := Low(WhereList) to High(WhereList) do
     if UpperCase(WhereList[j].Table) = UpperCase(TableName) then
     begin
        result := UpperCase(WhereList[j].Where);
        break;
     end;
  end;

begin
  result := true;
  CreateTT(ExpCaption, 'Export des tables en cours...', Length(TableList), true, true);

  for i := Low(TableList) to High(TableList) do
  try
    SetTTSubText('Ouverture de la table "'+TableToLibelle(TableList[i])+'"');
    StWhere := GetWhere(TableList[i]);
    if StWhere <> '' then StWhere := ' WHERE '+StWhere;
// Modif BBI 09/04 suite à nouveaux Tarifs
    if TableList[i] = 'TARIF' then DemoulineTarif;

    Q := OpenSQL('SELECT * FROM '+TableList[i]+StWhere, true);

    QueryToFormat(Q, 'Export_'+TableList[i], ParMSMQ, Fmt);
    Ferme(Q);

    SetTTValue(i);
    if TTCanceled then break;

  except
    on E : Exception do
     if not Quiet then PGIInfo('Erreur à l''exportation de la table "'+TableToLibelle(TableList[i])+'"'#13+E.Message, ExpCaption)
                  else raise;
  end;

  if TTCanceled then result := false;
  FreeTT;
end;

// Importe une ou plusieurs tables depuis des fichiers ou depuis MSMQ
function TETransfert.ImportFullTables(TableList : Array of String; ParMSMQ : boolean; Fmt : TTOBFormat) : boolean;
var i : integer;
begin
  result := true;
  CreateTT(ImpCaption, 'Import des tables en cours...', Length(TableList), true, true);

  for i := Low(TableList) to High(TableList) do
  try
    FormatToTable('Import_'+TableList[i], true, ParMSMQ, Fmt);

    SetTTValue(i);
    if TTCanceled then break;

  except
    on E : Exception do PGIInfo('Erreur à l''importation de la table "'+TableToLibelle(TableList[i])+'"'#13+E.Message, ImpCaption);
  end;

  if TTCanceled then result := false;
  FreeTT;
end;


procedure TETransfert.CreateTT(Caption, Titre : string; MaxValue : integer; EnableCancel, DisabledMenu : boolean);
begin
  if not Quiet then
  begin InitMoveProgressForm(Application.MainForm, Caption, Titre, MaxValue, EnableCancel, DisabledMenu);
        FTTAssigned := true; end;
end;

procedure TETransfert.SetTTMaxValue(Max : integer);
begin
end;

procedure TETransfert.SetTTValue(Val : integer);
begin
  if FTTAssigned then FTTCanceled := MoveCurProgressForm(FTTSubText);
end;

procedure TETransfert.SetTTSubText(St : String);
begin
  if FTTAssigned then FTTSubText := St;
end;

function TETransfert.TTCanceled : boolean;
begin
  result := false;
  if FTTAssigned then result := FTTCanceled;
end;

procedure TETransfert.FreeTT;
begin
  if FTTAssigned then begin FiniMoveProgressForm;
                            FTTAssigned := false; end;
end;
{$ENDIF}


function AddBkSlash(S : String) : String;
begin
  result := S;
  if S = '' then exit;
  if S[Length(S)] <> '\' then result := result + '\';
end;

function GetParamSocDef(Nom : String; Defaut : Variant) : Variant;
begin
  try
    result := GetParamSoc(Nom);
  except
    on exception do result := Defaut;
  end;
end;

procedure GetParams;
begin
  GotToZip := GetParamSocDef('SO_COMPRESS', GotToZip);
  ImpFolder := AddBkSlash(GetParamSocDef('SO_IMPFOLDER', ImpFolder));
  ExpFolder := AddBkSlash(GetParamSocDef('SO_EXPFOLDER', ExpFolder));
  Encrypted := GetParamSocDef('SO_GCENCRYPTEDMSG', Encrypted);
  InQueuePath := GetParamSocDef('SO_MSMQPATHIN', InQueuePath);
  OutQueuePath := GetParamSocDef('SO_MSMQPATHOUT', OutQueuePath);
  LiensOLEQuandTiers := GetParamSocDef('SO_GCESENDLO', LiensOLEQuandTiers);
  DoitNbSpizer := GetParamSocDef('SO_GCENBSP', DoitNbSpizer);
  ValidationAuto := GetParamSocDef('SO_GCEAUTOVALID', ValidationAuto);
end;
(*
function GetModeRegl(MR : String) : String;
begin
   result := GetParamSoc('SO_GCEMR'+MR);
//   if result = '' then result := GetParamSoc('SO_GCEMRCHQ');
end;
*)
function IsModeReglChq(GPMR : String) : boolean;
var i : integer;
    Q : TQuery;
    MPList : String;
begin
   Q := OpenSQL('SELECT MR_MP1, MR_MP2, MR_MP3, MR_MP4, MR_MP5, MR_MP6, MR_MP7, MR_MP8, MR_MP9, MR_MP10, MR_MP11, MR_MP12 FROM MODEREGL WHERE MR_MODEREGLE="'+GPMR+'"', true);
   MPList := '';
   for i := 1 to 12 do begin MPList := MPList+'"'+Q.{$IFDEF EAGLCLIENT}FindField{$ELSE}FieldByName{$ENDIF}('MR_MP'+inttostr(i)).AsString+'"';
                             if i < 12 then MPList := MPList + ', '; end;
   Ferme(Q);
   Q := OpenSQL('SELECT MP_LETTRECHEQUE FROM MODEPAIE WHERE MP_LETTRECHEQUE="X" AND MP_MODEPAIE IN ('+MPList+')', true);
   result := ((not Q.EOF) and (Q.{$IFDEF EAGLCLIENT}FindField{$ELSE}FieldByName{$ENDIF}('MP_LETTRECHEQUE').AsString = 'X'));
   Ferme(Q);
end;

function ModeReglAuto(MR : String) : boolean;
var Q : TQuery;
begin
   result := false;
   if ChampLogiqueExiste('MODEREGL', 'MR_EINTEGREAUTO') then
   begin
       Q := OpenSQL('SELECT MR_EINTEGREAUTO FROM MODEREGL WHERE MR_MODEREGLE="'+MR+'"', true);
       if not Q.EOF then result := (Q.{$IFDEF EAGLCLIENT}FindField{$ELSE}FieldByName{$ENDIF}('MR_EINTEGREAUTO').AsString = 'X');
       Ferme(Q);
   end
    else LogDebugMsg('Champ Logique n''existe pas !! MR_EINTEGREAUTO', true);

//   result := GetParamSoc('SO_GCERA'+MR);
end;

// Convertit une ETable en Table ; Rajoute les valeurs absentes dans la ETable au passage
procedure ETableToTable(SrcTOB, DestTOB : TOB);
var i_rec, i_fld : integer;
    Src, Dest, tcompl : TOB;
// tcompl : modif TABLERONDE
    NomChampSansPrefix,
    SrcPref, DestPref : String;
    Nat, Sou, Nb, Ind : String;
    Q : TQuery;
begin
  for i_rec := 0 to SrcTOB.Detail.Count-1 do
  begin
    Src := SrcTOB.Detail[i_rec];
    Dest := TOB.Create(Copy(Src.NomTable, 2, Length(Src.NomTable)), DestTOB, -1);

    SrcPref := TableToPrefixe(Src.NomTable);
    DestPref := TableToPrefixe(Dest.NomTable);

    for i_fld := 1 to Src.NbChamps do
    begin
      NomChampSansPrefix := Src.GetNomChamp(i_fld);
      Delete(NomChampSansPrefix, 1, Pos('_', NomChampSansPrefix)-1);
      if Dest.FieldExists(DestPref+NomChampSansPrefix) then
         Dest.PutValue(DestPref+NomChampSansPrefix, Src.GetValue(SrcPref+NomChampSansPrefix));
    end;

    if Dest.NomTable = 'PIECE' then
    begin
      Dest.PutValue('GP_VENTEACHAT', 'VEN');
//      if Not (ctxAffaire in V_PGI.PGIContexte) then
//         Dest.PutValue('GP_MODEREGLE', GetModeRegl(Src.GetValue('EP_MODEREGLE')));
      Dest.PutValue('GP_CBDATEEXPIRE', Src.GetValue('EP_CBDATEEXPIRE')); // PCS 21/03/2003
//      Dest.PutValue('GP_CREEPAR', 'IEC');
      Dest.PutValue('GP_ETATEXPORT', 'ATT');
      Dest.PutValue('GP_ETATCOMPTA', 'ATT');

//      if not ChampLogiqueExiste('PIECE', 'GP_RECALCULER') then Dest.AddChampSup('GP_RECALCULER',false);
//      Dest.PutValue('GP_RECALCULER', 'X');

      Q := OpenSQL('SELECT T_REPRESENTANT FROM TIERS WHERE T_TIERS="'+Src.GetValue('EP_TIERS')+'"', true);
      if not Q.EOF then
      try Dest.PutValue('GP_REPRESENTANT', Q.FindField('T_REPRESENTANT').AsVariant);
      except
          raise Exception.Create('Champ inconnu dans récupération infos représentant (pièce)');
      end;
      Ferme(Q);

      Dest.PutValue('GP_HEURECREATION', Now);
      Nat := Src.GetValue('EP_NATUREPIECEG');
      Sou := Src.GetValue('EP_SOUCHE');
      Nb := Src.GetValue('EP_NUMERO');
      Ind := Src.GetValue('EP_INDICEG');
      Q := OpenSQL('SELECT distinct EL_DOMAINE FROM ELIGNE WHERE EL_TYPELIGNE="ART" AND EL_NATUREPIECEG="'+Nat+'" AND EL_SOUCHE="'+Sou+'" AND EL_NUMERO='+Nb+' AND EL_INDICEG='+Ind, true);
      if not Q.EOF then
      try Dest.PutValue('GP_DOMAINE', Q.FindField('EL_DOMAINE').AsVariant);
      except
          raise Exception.Create('Champ inconnu dans récupération infos domaine (pièce)');
      end;
      Ferme(Q);
    end;

    if Dest.NomTable = 'LIGNE' then
    begin
      Dest.PutValue('GL_REFARTSAISIE', Src.GetValue('EL_CODEARTICLE'));
      Dest.PutValue('GL_PERIODE', GetPeriode(Src.GetValue('EL_DATEPIECE')));
      Dest.PutValue('GL_SEMAINE', NumSemaine(Src.GetValue('EL_DATEPIECE')));
      //Dest.PutValue('GL_DEPOT', VH_GC.GCDepotDefaut);
      Dest.PutValue('GL_TYPEDIM', 'NOR');

      Q := OpenSQL('SELECT GA_DPA, GA_PMAP, GA_PMRP, GA_DPR, GA_FAMILLENIV1, GA_FAMILLENIV2, GA_FAMILLENIV3, GA_LIBREART1, '+
                   'GA_LIBREART2, GA_LIBREART3, GA_LIBREART4, GA_LIBREART5, GA_LIBREART6, GA_LIBREART7, GA_LIBREART8, GA_LIBREART9, GA_LIBREARTA '+
                   'FROM ARTICLE WHERE GA_ARTICLE="'+Src.GetValue('EL_ARTICLE')+'"', true);
      if not Q.EOF then
      try Dest.PutValue('GL_DPA', Q.FindField('GA_DPA').AsVariant);
          Dest.PutValue('GL_PMAP', Q.FindField('GA_PMAP').AsVariant);
          Dest.PutValue('GL_PMRP', Q.FindField('GA_PMRP').AsVariant);
          Dest.PutValue('GL_PMAPACTU', Q.FindField('GA_PMAP').AsVariant);
          Dest.PutValue('GL_PMRPACTU', Q.FindField('GA_PMRP').AsVariant);
          Dest.PutValue('GL_DPR', Q.FindField('GA_DPR').AsVariant);
          Dest.PutValue('GL_FAMILLENIV1', Q.FindField('GA_FAMILLENIV1').AsVariant);
          Dest.PutValue('GL_FAMILLENIV2', Q.FindField('GA_FAMILLENIV2').AsVariant);
          Dest.PutValue('GL_FAMILLENIV3', Q.FindField('GA_FAMILLENIV3').AsVariant);
          Dest.PutValue('GL_LIBREART1', Q.FindField('GA_LIBREART1').AsVariant);
          Dest.PutValue('GL_LIBREART2', Q.FindField('GA_LIBREART2').AsVariant);
          Dest.PutValue('GL_LIBREART3', Q.FindField('GA_LIBREART3').AsVariant);
          Dest.PutValue('GL_LIBREART4', Q.FindField('GA_LIBREART4').AsVariant);
          Dest.PutValue('GL_LIBREART5', Q.FindField('GA_LIBREART5').AsVariant);
          Dest.PutValue('GL_LIBREART6', Q.FindField('GA_LIBREART6').AsVariant);
          Dest.PutValue('GL_LIBREART7', Q.FindField('GA_LIBREART7').AsVariant);
          Dest.PutValue('GL_LIBREART8', Q.FindField('GA_LIBREART8').AsVariant);
          Dest.PutValue('GL_LIBREART9', Q.FindField('GA_LIBREART9').AsVariant);
          Dest.PutValue('GL_LIBREARTA', Q.FindField('GA_LIBREARTA').AsVariant);
      except
          raise Exception.Create('Champ inconnu dans récupération infos article');
      end;
      Ferme(Q);

//      if not ChampLogiqueExiste('LIGNE', 'GL_RECALCULER') then Dest.AddChampSup('GL_RECALCULER',false);
//      Dest.PutValue('GL_RECALCULER', 'X');

      Q := OpenSQL('SELECT T_REPRESENTANT FROM TIERS WHERE T_TIERS="'+Src.GetValue('EL_TIERS')+'"', true);
      if not Q.EOF then
      try Dest.PutValue('GL_REPRESENTANT', Q.FindField('T_REPRESENTANT').AsVariant);
      except
          raise Exception.Create('Champ inconnu dans récupération infos représentant (ligne)');
      end;
      Ferme(Q);

      if Src.GetValue('EL_TYPELIGNE')='ART' then Dest.PutValue('GL_TYPEREF', 'ART');
      Dest.PutValue('GL_SOCIETE', V_PGI.CodeSociete);
    end;

    if Dest.NomTable = 'TIERS' then
    begin
      Dest.PutValue('T_LETTRABLE','X');
      Dest.PutValue('T_SOLDEPROGRESSIF','X');
      Dest.PutValue('T_SAUTPAGE','-');
      Dest.PutValue('T_TOTAUXMENSUELS','-');
      Dest.PutValue('T_SOCIETE', V_PGI.CodeSociete);
      if (Dest.GetValue('T_NATUREAUXI') = null) or (Dest.GetValue('T_NATUREAUXI') = '')
         then Dest.PutValue('T_NATUREAUXI', 'CLI');

// modif TABLERONDE
      tcompl := TOB.Create('TIERSCOMPL', DestTOB, -1);
      tcompl.PutValue('YTC_AUXILIAIRE',Dest.GetValue('T_AUXILIAIRE'));
      tcompl.PutValue('YTC_TIERS',Dest.GetValue('T_TIERS'));
{$IFDEF TABLERONDE}
      tcompl.PutValue('YTC_TABLELIBRETIERS1','04');
      tcompl.PutValue('YTC_TABLELIBRETIERS2','I');
{$ENDIF TABLERONDE}
// fin modifs TABLERONDE

//      if Not (ctxAffaire in V_PGI.PGIContexte) then
//         Dest.PutValue('T_MODEREGLE', GetModeRegl(Src.GetValue('ETI_MODEREGLE')));
    end;
  end;
end;

const MyEditIsRich : TRichEdit = nil;

function DeRTFize(S : String; ParentForm : TForm = nil) : String;
var Str : TStringStream;
begin
   {$IFDEF SERVICE}
   if ParentForm = nil then raise Exception.Create('Pas de fenêtre parent pour le déRTFizator');
   {$ELSE}
   if ParentForm = nil then ParentForm := Application.MainForm;
   {$ENDIF}
   if MyEditIsRich = nil then
   begin
      MyEditIsRich := TRichEdit.Create(ParentForm);
      with MyEditIsRich do
      begin
        ParentWindow := ParentForm.Handle;
        Visible := false;
        WordWrap := false;
      end;
   end;

   with MyEditIsRich do
   begin
      Str := TStringStream.Create(S);
      Str.Seek(0, soFromBeginning);
      Lines.LoadFromStream(Str);
      Str.Free;

      result := Text;
   end;
end;

procedure CloseDeRTFizator;
begin
   if MyEditIsRich = nil then exit;
   MyEditIsRich.Free;
   MyEditIsRich := nil;
end;

function NbSpize(S : String) : String;
var i : integer;
begin
   if DoitNbSpizer then
   begin
      result := StringReplace(S, ' ', '&nbsp;', [rfReplaceAll]);
      //result := StringReplace(result, #13, '<BR>'#13, [rfReplaceAll]);
      for i := Length(result) downto 1 do
       if result[i] = #13 then Insert('<BR>', result, i);
   end
    else result := S;
end;

// Convertit les espaces en &nbsp; au passage
procedure DeRTFizeLiensOLE(T : TOB);
var i : integer;
begin
   for i := 0 to T.Detail.Count-1 do
   begin
    if T.Detail[i].GetValue('LO_QUALIFIANTBLOB') = 'MEM' then
     T.Detail[i].PutValue('LO_OBJET', NbSpize(DeRTFize(T.Detail[i].GetValue('LO_OBJET'))));
   end;
   CloseDeRTFizator;
end;

// Convertit une Table en ETable
procedure TableToETable(SrcTOB, DestTOB : TOB);
var i_rec, i_fld : integer;
    Src, Dest : TOB;
    NomChampSansPrefix,
    SrcPref, DestPref : String;
begin
  for i_rec := 0 to SrcTOB.Detail.Count-1 do
  begin
    Src := SrcTOB.Detail[i_rec];
    Dest := TOB.Create('E'+Src.NomTable, DestTOB, -1);

    SrcPref := TableToPrefixe(Src.NomTable);
    DestPref := TableToPrefixe(Dest.NomTable);

    for i_fld := 1 to Dest.NbChamps do
    begin
      NomChampSansPrefix := Dest.GetNomChamp(i_fld);
      if uppercase(Copy(NomChampSansPrefix,Length(NomChampSansPrefix)-2,3))='CON' then Continue;
      Delete(NomChampSansPrefix, 1, Pos('_', NomChampSansPrefix)-1);
      if ChampType(SrcPref+NomChampSansPrefix) = 'BLOB'
        then Dest.PutValue(DestPref+NomChampSansPrefix, DeRTFize(Src.GetValue(SrcPref+NomChampSansPrefix)))
        else Dest.PutValue(DestPref+NomChampSansPrefix, Src.GetValue(SrcPref+NomChampSansPrefix));
    end;
  end;
  CloseDeRTFizator;
end;

procedure DemoulineTarif;
var TobYTarifs, TobYTarifsFour : TOB;
    TobTarif, TobLigTarif : TOB;
    TobLigYTarifs, TobLigYTarifsFour : TOB;
    TobLoc : TOB;
    ind1, ind2, Rang : integer;
    Q : TQuery;
    Remise : string;
begin
//  vu qu'on a rien à faire de la table précedente, on la purge purement et simplement...
  ExecuteSQL('Delete from TARIF');
//  ceci fait, on mouline les tables YTARIFS et YTARIFSFOURCHETTE pour regenerer la table
//  TARIF, Traitement inverse de celui fait au passage en V6... (travail d'arabe tout ça...)
//  surtout que au passage on perd la notion de tarif global !!!!
  Q := OpenSQL('SELECT * FROM YTARIFS', true);
  if Q.EOF then
  begin
    Ferme(Q);
    Exit;
  end;
  TobYTarifs := TOB.Create('Les Tarifs', nil, -1);
  TobYTarifs.LoadDetailDB('YTARIFS', '', '', Q, False);

  for ind1 := 0 to TobYTarifs.Detail.Count - 1 do
  begin
    TobLigYTarifs := TobYTarifs.Detail[ind1];
    Q := OpenSQL('SELECT * FROM YTARIFSFOURCHETTE where YTF_IDENTIFIANTYTS=' +
                 IntToStr(TobLigYTarifs.GetValue('YTS_IDENTIFIANT')), true);
    if not Q.EOF then
      TobLigYTarifs.LoadDetailDB('YTARIFSFOURCHETTE', '', '', Q, False);
    Ferme(Q);
  end;

  Rang := 1;
  TobTarif := TOB.Create('Les Tarifs Ancienne Mode', nil, -1);
  for ind1 := 0 to TobYTarifs.Detail.Count - 1 do
  begin
    TobLigYTarifs := TobYTarifs.Detail[ind1];
    TobLigTarif := TOB.Create('TARIF', TobTarif, -1);
    TobLigTarif.InitValeurs;
    TobLigTarif.PutValue('GF_TARIF', Rang);
    Inc(Rang);
    if TobLigYTarifs.GetValue('YTS_FONCTIONNALITE') = sTarifClient then
      TobLigTarif.PutValue('GF_NATUREAUXI', 'CLI')
    else if TobLigYTarifs.GetValue('YTS_FONCTIONNALITE') = sTarifFournisseur then
      TobLigTarif.PutValue('GF_NATUREAUXI', 'FOU');
    TobLigTarif.PutValue('GF_DEPOT' , TobLigYTarifs.GetValue('YTS_DEPOT'));
    TobLigTarif.PutValue('GF_TIERS' , TobLigYTarifs.GetValue('YTS_TIERS'));
    TobLigTarif.PutValue('GF_TARIFTIERS' , TobLigYTarifs.GetValue('YTS_TARIFTIERS'));
    TobLigTarif.PutValue('GF_DEVISE' , TobLigYTarifs.GetValue('YTS_DEVISE'));
    TobLigTarif.PutValue('GF_DATEDEBUT' , TobLigYTarifs.GetValue('YTS_DATEDEBUT'));
    TobLigTarif.PutValue('GF_DATEFIN' , TobLigYTarifs.GetValue('YTS_DATEFIN'));
    TobLigTarif.PutValue('GF_ARTICLE' , TobLigYTarifs.GetValue('YTS_ARTICLE'));
    TobLigTarif.PutValue('GF_TARIFARTICLE' , TobLigYTarifs.GetValue('YTS_TARIFARTICLE'));
    TobLigTarif.PutValue('GF_DATECREATION' , TobLigYTarifs.GetValue('YTS_DATECREATION'));
    TobLigTarif.PutValue('GF_DATEMODIF' , TobLigYTarifs.GetValue('YTS_DATEMODIF'));
    TobLigTarif.PutValue('GF_LIBELLE' , TobLigYTarifs.GetValue('YTS_LIBELLETARIF'));
    if TobLigYTarifs.GetValue('YTS_ACTIF') = 'X' then
      TobLigTarif.PutValue('GF_FERME' , '-')
    else
      TobLigTarif.PutValue('GF_FERME' , 'X');
    if TobLigYTarifs.GetValue('YTS_TTCOUHT') = 'X' then
      TobLigTarif.PutValue('GF_REGIMEPRIX', 'TTC')
    else if TobLigYTarifs.GetValue('YTS_TTCOUHT') = '-' then
      TobLigTarif.PutValue('GF_REGIMEPRIX', 'HT');
    if TobLigYTarifs.GetValue('YTS_CASCREMISETIER') = 'MEI' then
      TobLigTarif.PutValue('GF_CASCADEREMISE', 'MIE')
    else if TobLigYTarifs.GetValue('YTS_CASCREMISETIER') = 'ANN' then
    begin
      if TobLigYTarifs.GetValue('YTS_CASCCONTEXTE') = 'X' then
        TobLigTarif.PutValue('GF_CASCADEREMISE', 'CAS')
      else
        TobLigTarif.PutValue('GF_CASCADEREMISE', 'FOR');
    end
    else if TobLigYTarifs.GetValue('YTS_CASCREMISETIER') = 'CUM' then
      TobLigTarif.PutValue('GF_CASCADEREMISE', 'CUM');
    if TobLigYTarifs.Detail.Count = 0 then
    begin
      TobLigTarif.PutValue('GF_PRIXUNITAIRE' , TobLigYTarifs.GetValue('YTS_PRIXBRUT'));
      Remise := '';
      if TobLigYTarifs.GetValue('YTS_REMISE1') <> 0 then
        Remise := FloatToStr(TobLigYTarifs.GetValue('YTS_REMISE1'));
      if TobLigYTarifs.GetValue('YTS_REMISE2') <> 0 then
        Remise := Remise + '+' + FloatToStr(TobLigYTarifs.GetValue('YTS_REMISE2'));
      if TobLigYTarifs.GetValue('YTS_REMISE2') <> 0 then
        Remise := Remise + '+' + FloatToStr(TobLigYTarifs.GetValue('YTS_REMISE3'));
      TobLigTarif.PutValue('GF_CALCULREMISE' , Remise);
      TobLigTarif.PutValue('GF_REMISE', RemiseResultante(Remise));
      TobLigTarif.PutValue('GF_QUANTITATIF' , '-');
      TobLigTarif.PutValue('GF_BORNEINF' , -999999);
      TobLigTarif.PutValue('GF_BORNESUP' , 999999);
      CalcPriorite(TobLigTarif);
    end
    else
    begin
      for ind2 := 0 to TobLigYTarifs.Detail.Count - 1 do
      begin
        TobLigYTarifsFour := TobLigYTarifs.Detail[ind2];
        if ind2 > 0 then
        begin
          TobLoc := TOB.Create('TARIF', TobTarif, -1);
          TobLoc.Dupliquer(TobLigTarif, True, True, True);
          TobLigTarif := TobLoc;
          TobLigTarif.PutValue('GF_TARIF', Rang);
          Inc(Rang);
        end;
        TobLigTarif.PutValue('GF_LIBELLE' , TobLigYTarifsFour.GetValue('YTF_LIBELLETARIF'));
        TobLigTarif.PutValue('GF_PRIXUNITAIRE' , TobLigYTarifsFour.GetValue('YTF_PRIXBRUT'));
        Remise := '';
        if TobLigYTarifs.GetValue('YTS_REMISE1') <> 0 then
          Remise := FloatToStr(TobLigYTarifs.GetValue('YTS_REMISE1'));
        if TobLigYTarifs.GetValue('YTS_REMISE2') <> 0 then
          Remise := Remise + '+' + FloatToStr(TobLigYTarifs.GetValue('YTS_REMISE2'));
        if TobLigYTarifs.GetValue('YTS_REMISE2') <> 0 then
          Remise := Remise + '+' + FloatToStr(TobLigYTarifs.GetValue('YTS_REMISE3'));
        TobLigTarif.PutValue('GF_CALCULREMISE' , Remise);
        TobLigTarif.PutValue('GF_REMISE', RemiseResultante(Remise));
        TobLigTarif.PutValue('GF_QUANTITATIF' , 'X');
        TobLigTarif.PutValue('GF_BORNEINF' , TobLigYTarifsFour.GetValue('YTF_FOURCHETTEMINI'));
        TobLigTarif.PutValue('GF_BORNESUP' , TobLigYTarifsFour.GetValue('YTF_FOURCHETTEMAXI'));
        if TobLigYTarifsFour.GetValue('YTF_FOURCHETTEMAXI') = n9999999999 then
          TobLigTarif.PutValue('GF_BORNESUP' , 999999)
        else
          TobLigTarif.PutValue('GF_BORNESUP' , TobLigTarif.GetValue('GF_BORNESUP') - 1);
        CalcPriorite(TobLigTarif);
      end;
    end;
  end;
  TobTarif.InsertOrUpdateDB();
end;

end.
