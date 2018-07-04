unit UDataMemAGL;

interface

uses Classes,
  ULList, UEnreg;

const
  maxDatamem = 85;
  maxNbRefCle = 40;

  dmErr_ok = 0;
  dmErr_finEnreg = 1;
  dmErr_nontrouve = 4;
  dmErr_suppNonDef = 5;
  dmErr_erreur = 100;
  dmErr_dejaExistant = 101;
  dmErr_nondef = 102; // n° de datamem non créé
  dmErr_keyViol = 103; // n° de datamem non créé

type
  noDM = word;

  noRefCle = word;

  TDataMem = class(Tobject)
  private
    buff: Tenreg;
    function Rech_recurs(const ListeChpCle: array of string; deb, fin: integer; var IndInsert: integer): integer;
    function okCle(LaCase: TLList; const ListeChpCle: array of string): integer;
    procedure genList(LaCase: TLList; Enr: TstringList);
    procedure genCledm(Lacase: TLList; const TabEnr: array of string);
    function donneRef(IndCle: integer; const S: string): PChar;
    procedure rempliBuff(Curs: integer);
    procedure clear;
    procedure addSql(const Sql: string);
  public
    compteurPseq: integer;
    cleUnique: boolean;
    // Numéro de reconnaissance du DataMem
    NumDM: noDM;
    // Nombre de clé
    NbCle: word;
    // Type des champs
    TypeChamp: TstringList;
    // Curseur en cours
    CursDm: integer;
    // Liste des cases
    ListCase: Tlist;
    // Liste de Listes correspondant aux clés
    ListnoRefCle: array[0..12] of noRefCle;
    // Test si la compression de données doit être activée
    FokCompress: boolean;
    constructor create(NombreChCle: word; OkCompress: word; OkCleunique: boolean);
    destructor destroy(); override;
    //Utilisé dans le cas ou le remplissage se fait sans Sql.
    //Attention : le remplissage doit se faire dans l'ordre de tri des clés
    //( différent de dm_creaEnreg )
    //Retour : 0 : OK  , >= 100 : erreur
    function add(const TabEnr: array of string): integer;

  end;

  TdmPseq = class(Tobject)
  public
    DM: TDatamem;
    debut: boolean; // le curseur n'avance pas au 1er appel
    curseur: integer;
    filtre: TStringList;
    yaFiltre: boolean;
    constructor create(LeDM: Tdatamem; const ListeChp: array of string; okFiltre: boolean);
    destructor destroy(); override;
  end;

procedure initUdatamem;

function dm_sqlInit(NomDM: noDM; OkCompress: word; const Sql: string;
  const TypeToutChamp: array of string; OKCleUnique: boolean;
  const ListNoRef: array of norefcle): integer;
//NomDM : n°du DATAMEM qui fera référence pour la manipulation des données
//Sera une constante en idm_
// Sql : commande sql qui lit les données ATTENTION !! Il faut un ORDER BY
//NbCle : Nb de champ clé du DATAMEM,
//Retour : 0 : OK  , >=100 : erreur

function dm_addSql(NomDM: noDM; const Sql: string): integer;
// rempli un datamem déja crée
// Attention,il y a un clear avant le remplissage par le sql

// Initialise le DataMem
function dm_init(NomDM: noDM; OkCompress: word; const TypeToutChamp: array of string;
  OKCleUnique: boolean; const ListNoRef: array of norefcle;
  var DataMem: Tdatamem): integer;
//DataMem : pointer sur l'objet DATAMEM
//Retour : 0 : OK  , >= 100 : erreur

// Enlève le DataMem
function dm_supprime(NomDM: noDM): integer;
//Supprime le DATAMEM.
//Retour : 0 : OK  , >= 100 : erreur

// vide le contenu du datamem
function dm_clear(NomDM: noDM): integer;

// Recherche l'existence d'un DataMem
function RechercheDM(NomDM: noDM; var DM: TdataMem): boolean;

procedure destroyAllDm();

function dm_trouveEnr(NomDM: noDM; const ListeChpCle: array of string; var Enreg: Tenreg): integer;
//ListeChp doit contenir que les x premiers champ clés
//Place le curseur sur l'enreg
//Retour : 0: OK  ,  4: enreg non trouvé ,   >= 100 : erreur

function dm_pseq(NomDM: noDM; const ListeChp: array of string; var Pseq: TdmPseq): integer;
//Initialise une lecture séquentiel d'un DATAMEM
//Retour : 0 : OK  , >= 100 : erreur
//Initialise le curseur sur le 1er enreg recherché

// pour parcours complet
function dm_pseqAll(NomDM: noDM; var Pseq: TdmPseq): integer;

function dm_pseqLect(Pseq: TdmPseq; var Enreg: Tenreg): integer;
//Renvoie l'enreg sur le curseur et se déplace au suivant.
// le tdmpseq doit être obtenu via dm_pseq
//La fonction renvoie 1 si l'enreg courant ne correspond pas au filtre.
//L'enreg ne sera valide que si la fonction retourne 0.
//Retour : 0: OK  ,  1: Fin des enreg ,   >= 100 : erreur

// Attention si on sort de la boucle avant la fin du pseqlect, il faut faire
// un free sur le TdmPseq

// Bookmark
function dm_GetCurseur(NomDM: noDM): integer;
procedure dm_SetCurseur(NomDM: noDM; Curseur: integer);

// Renvoit l'enregistrement suivant s'il existe
function dm_suivant(NomDm: noDM; var Enreg: Tenreg): integer;

// Renvoit le nombre d'enregistrement
function dm_Nb_Enregistrement(NomDM: noDM): integer;

function dm_SetEnreg(NomDm: noDM; var Enreg: TEnreg; Curseur: integer): integer;

//supprime l'enregistrement courant
function dm_suppEnregCour(NomDM: noDM): integer;

//modifie le champ de l'enreg en cours par Chp
function dm_modifchcour(NomDM: noDM; NumChp: noChamp; const Chp: string): integer;
//Travaille sur l'enregistrement courant
//NumChp correspond au numéro de champ de l'enregistrement DataMem

//modifie les champs  non en clé sur l'enreg courant
function dm_modifEnregCour(nomDM: noDM; const TabEnr: array of string): integer;

//supprime un enregistrement
function dm_suppEnreg(NomDM: noDM; const ListeChpCle: array of string): integer;
//Supprime l'enregistrement si il a été trouvé
//Retourne : 0 pour OK, 4 pour non trouvé, 100>erreur

  //ajoute un enregistrement
function dm_creaEnreg(NomDM: noDM; const TabEnr: array of string): integer;
//ajoute l'enreg à sa place de façon triée

{$IFDEF TRACE}
procedure dm_debug(nomDM: noDM);
// édite le contenu d'un datamem dans le debugLog
{$ENDIF}

function dmErrLibelle(Err: integer): string;

procedure freeListRef(L: TList);

function Recherche_Code_ListRef_i(Liste: TList; const S: string; deb, fin: integer; var Pos: integer): boolean;

var ListDataMem: array[1..maxDatamem] of Tdatamem;
  LaListRef: array[0..maxNbRefCle] of TList;
  ListeNil: array[0..0] of string;
  //    BZIP: Tbzip2;
 //     dm:TDataMem;

implementation

uses sysUtils, Dialogs,
  {$IFNDEF EAGLCLIENT}
  Db, DbTables,
  {$ENDIF}
  HCtrls, UTob, Hent1, HMsgBox,
  {$IFDEF TRACE}
  hdebug,
  {$ENDIF}
  DMUtil, uutil;

constructor Tdatamem.create(NombreChCle: word; OkCompress: word; OkCleunique: boolean);
begin
  try
    inherited create();
    compteurPseq := 0;
    FOKCompress := true;
    if OkCompress = 0 then
      FOKCompress := false;
    Nbcle := NombreChCle;
    cleUnique := OkCleUnique;
    ListCase := TList.create;
    // ListRef:=TList.create;
   //  for I:=1 to NbCle do
   //    ListRef.add(TList.create);
    buff := Tenreg.create;
    TypeChamp := TStringList.create;

  except
    on E: Exception do
    begin
      raise(exception.create('Pb dm_create ' + E.Message));
    end;
  end;
end;

destructor Tdatamem.destroy();
begin
  try
    clear;
    listcase.free;
    //   with ListRef do
    //    begin
    //      for j:= 0 to (count -1) do
    //         TList(items[j]).free;
    //      free;
    //    end;
    buff.free;
    TypeChamp.free;
    if compteurPseq <> 0
      then HShowmessage('1;Attention;Pas de libération tDmPseq ( %% ) dm N°$$ ;W;O;O;O', IntToStr(compteurPseq), IntToStr(numDM));
    inherited destroy;
  except
    on E: Exception do
    begin
      raise(exception.create('Pb dm_destroy ' + E.Message));
    end;
  end;
end;

procedure Tdatamem.clear();
var C: TLList;
  j: integer;
begin
  try
    with listcase do
    begin
      for j := 0 to (count - 1) do
      begin
        C := items[j];
        if FOkCompress
          then TMemoryStream(C.items[0]).free //Tstream libéré
        else strDispose(PChar(C.items[0])); //Pchar libéré
        C.free; // TLList libéré, les Pchar de cette liste sont libérés via laListref
      end;
      clear;
    end;
    { on touche pas à la Listref lors d'un clear
   with ListRef do
     begin
       for j:= 0 to (count -1) do
         begin
           for k:= 0 to (TList(items[j]).count -1) do
             begin
               S:=TList(items[j]).items[k];
               strDispose(S);
             end;
          TList(items[j]).clear;
         end;
     end;
   }
  except
    on E: Exception do
    begin
      raise(exception.create('Pb clear' + E.Message));
    end;
  end;
end;

constructor TdmPseq.create(LeDM: Tdatamem; const ListeChp: array of string; okFiltre: boolean);
var i: integer;
begin

  try
    inherited create();
    debut := true;
    dm := LeDM;
    curseur := LeDM.cursdm;
    inc(LeDM.compteurPseq); // Pour compter le nombre de pseq en cours
    yaFiltre := okfiltre;
    Filtre := TstringList.create;
    if yafiltre then
      for i := 0 to high(ListeChp) do
        Filtre.add(ListeChp[i]);
  except
    on E: Exception do
    begin
      raise(exception.create('Pb dm_create ' + E.Message));
    end;
  end;

end;

destructor TdmPseq.destroy();
begin
  dec(DM.compteurPseq);
  Filtre.free;
  inherited destroy;
end;

procedure destroyAllDm();
var I: integer;
  DM: Tdatamem;
  LR: TList;
begin
  for I := 1 to maxDatamem do
  begin
    DM := ListDataMem[I];
    if assigned(DM) then DM.free;
    ListDataMem[I] := nil;
  end;
  for I := 0 to maxNbRefcle do
  begin
    LR := laListRef[I];
    if assigned(LR) then freeListRef(LR);
    LaListRef[I] := nil;
  end;
end;

procedure freeListRef(L: TList);
var k: integer;
  S: PChar;
begin
  with L do
  begin
    for k := 0 to (count - 1) do
    begin
      S := items[k];
      strDispose(S);
    end;
    free;
  end;
end;

{
procedure TdataMem.Trier(Liste:TList);
var
cont : boolean;
i,nb:integer;
Profil_inter: TProfil;
begin
Cont := true;
nb := Liste.Count-1;
Profil_inter := TProfil.Create;
while cont do
begin
  cont := false;
  for i := 0 to nb -1 do
      if TPROFIL(Liste[i]).heure_debut > TPROFIL(Liste[i+1]).heure_debut
        then
          begin
            Copier_Profil(TPROFIL(Liste[i]).heure_debut,TPROFIL(Liste[i]).heure_fin,TPROFIL(Liste[i]).potentiel,TPROFIL(Liste[i]).version,Profil_inter);
            Copier_Profil(TPROFIL(Liste[i+1]).heure_debut,TPROFIL(Liste[i+1]).heure_fin,TPROFIL(Liste[i+1]).potentiel,TPROFIL(Liste[i+1]).version,TPROFIL(Liste[i]));
            Copier_Profil(Profil_inter.heure_debut,Profil_inter.heure_fin,Profil_inter.potentiel,Profil_inter.version,TPROFIL(Liste[i+1]));
            cont := True;
          end;
end;
Profil_inter.Free;
end;
}

function TdataMem.add(const TabEnr: array of string): integer;
var
  Lacase: TLList;
  S: string;
  // MS: TMemoryStream ;
  // SSt: TStringStream ;
begin
  try
    Lacase := TLList.create(nbcle + 1);
    listCase.add(Lacase);
    cursdm := (listcase.count - 1); // car l'ajout est en fin
    S := genenreg(Nbcle, TabEnr, #1);
    {
    if FOkCompress
     then
      begin
        MS:=TMemoryStream.create;
        SSt:=TstringStream.create(S);
        BZIP.compressStream(SSt,MS);
        SSt.free;
        Lacase.items[0]:= MS; // pointe sur le block compressé
      end
     else
     }
    Lacase.items[0] := strnew(PChar(S)); // création du PChar à l'indice 0
    genCledm(Lacase, TabEnr);
    Result := dmErr_ok;
  except
    on E: Exception do
    begin
      //result:=100;
      raise(exception.create('Pb dm_add ' + E.Message));
    end;
  end;

end;

function RechercheDM(NomDM: noDM; var DM: TdataMem): boolean;
begin
  if NomDM = 0 then
  begin
    Result := false;
    exit;
  end;
  DM := ListDatamem[NomDM];
  if assigned(DM) then
    Result := true
  else
    Result := false;
end;

function dm_sqlInit(NomDM: noDM; OkCompress: word; const Sql: string;
  const TypeToutChamp: array of string; OKCleUnique: boolean;
  const ListNoRef: array of norefcle): integer;
var
  Init: integer;
  DataMem: TDataMem;
begin
  try
    Init := dm_Init(NomDm, OkCompress, TypeToutChamp, OkCleUnique, ListNoRef, DataMem);
    if Init = dmErr_ok
      then
    begin
      DataMem.addSql(Sql);
    end;
    Result := Init;
  except
    on E: Exception do
    begin
      //result:=100;
      raise(exception.create('Pb dm_sqlInit dm n° ' + IntToStr(NomDm) + ' ' + E.Message));
    end;
  end;
end;

function dm_init(NomDM: noDM; OkCompress: word; const TypeToutChamp: array of string;
  OKCleUnique: boolean; const ListNoRef: array of norefcle;
  var DataMem: Tdatamem): integer;
var I: integer;
  _: Tdatamem;
begin
  try
    if rechercheDM(NomDM, _) then
    begin
      DataMem := nil;
      Result := dmErr_dejaExistant;
      exit;
    end;
    DataMem := TdataMem.create(high(TypeToutChamp) + 1, OkCompress, OkCleUnique);
    DataMem.NumDM := NomDM;
    DataMem.cursDm := -1; // pointe sur rien
    ListDataMem[NomDm] := DataMem;
    if high(ListNoRef) <> high(TypeToutChamp) then raise(exception.create('liste refclé incorrecte'));
    for I := 0 to high(TypeToutChamp) do
    begin
      DataMem.TypeChamp.add(TypeToutChamp[I]);
      DataMem.ListNoRefCle[i] := ListNoRef[I];
    end;
    Result := dmErr_ok;
  except
    on E: Exception do
    begin
      //result:=100;
      raise(exception.create('Pb dm_Init dm n° ' + IntToStr(NomDm) + E.Message));
    end;
  end;
end;

function dm_supprime(NomDM: noDM): integer;
var DM: TdataMem;
begin
  try
    if RechercheDM(NomDM, DM)
      then
    begin
      DM.free;
      ListDataMem[NomDM] := nil;
      Result := dmErr_ok;
    end
    else Result := dmErr_Suppnondef;
  except
    on E: Exception do
    begin
      //result:=100;
      raise(exception.create('Pb dm_supprime dm n° ' + IntToStr(NomDm) + E.Message));
    end;
  end;
end;

function dm_clear(NomDM: noDM): integer;
var DM: TdataMem;
begin
  try
    if RechercheDM(NomDM, DM)
      then
    begin
      DM.clear;
      Result := dmErr_ok;
    end
    else Result := dmErr_nontrouve;
  except
    on E: Exception do
    begin
      //result:=100;
      raise(exception.create('Pb dm_clear dm n° ' + IntToStr(NomDm) + E.Message));
    end;
  end;
end;

function dm_addSql(NomDM: noDM; const Sql: string): integer;
var DM: TdataMem;
begin
  try
    if RechercheDM(NomDM, DM)
      then
    begin
      DM.clear;
      DM.addSql(Sql);
      Result := dmErr_ok;
    end
    else Result := dmErr_nontrouve;
  except
    on E: Exception do
    begin
      //result:=100;
      raise(exception.create('Pb dm_addSql dm n° ' + IntToStr(NomDm) + E.Message));
    end;
  end;
end;

function dm_trouveEnr(NomDM: noDM;
  const ListeChpCle: array of string; var Enreg: Tenreg): integer;
var DM: TdataMem;
  ii, Item: integer;
begin
  try
    if RechercheDM(NomDM, DM)
      then
    begin
      Item := DM.Rech_recurs(ListeChpCle, 0, DM.ListCase.Count - 1, ii);
      if Item <> -1
        then
      begin
        DM.cursdm := Item;
        DM.rempliBuff(Item); // lit la case
        Enreg := DM.buff;
        dm_trouveEnr := dmErr_ok;
      end
      else
      begin
        Enreg := nil;
        dm_trouveEnr := dmErr_nontrouve;
      end;
    end
    else dm_trouveEnr := dmErr_nondef;
  except
    on E: Exception do
    begin
      //result:=100;
      raise(exception.create('Pb dm_trouveenr dm n° ' + IntToStr(NomDm) + E.Message));
    end;
  end;
end;

function TDataMem.Rech_recurs(const ListeChpCle: array of string; deb, fin: integer; var IndInsert: integer): integer;
var Item: integer;
  OKC: integer;
begin
  try
    if deb > fin
      then
    begin
      Result := -1;
      IndInsert := deb; // c'est l'indice ou peux s'insérer l'enreg
      exit;
    end;
    Item := (deb + fin) div 2;
    OKC := okCle(ListCase.items[Item], ListeChpCle);
    if OKC = 0
      then
    begin

      if (Item <> 0) and ((nbCle > high(ListeChpCle)) or (not cleUnique))
        then
      begin // on recule pour se positionner sur le 1er enreg ok
        repeat
          Item := Item - 1;
          OKC := okCle(ListCase.items[Item], ListeChpCle);
        until (OKC <> 0) or (Item = 0);
        if OKC <> 0
          then Item := Item + 1;
      end;
      Result := Item;
      IndInsert := Item; // pour insertion avec clé non unique
    end
    else
      if OKC > 0
      then
      Result := Rech_Recurs(ListeChpCle, Item + 1, fin, IndInsert)
    else
      Result := Rech_Recurs(ListeChpCle, deb, Item - 1, IndInsert)
  except
    on E: Exception do
    begin
      //result:=100;
      raise(exception.create('Pb dm_rech_curs ' + E.Message));
    end;
  end;
end;

function dm_pseq(NomDM: noDM; const ListeChp: array of string; var Pseq: TdmPseq): integer;
var DM: TDataMem;
  _Enr: Tenreg;
begin
  try
    if RechercheDM(NomDM, DM)
      then
    begin
      if DM.listCase.count = 0 then
        Result := dmErr_nontrouve
      else // une hypothèse est que ListeChp a au moins 1 élément
        Result := dm_trouveEnr(NomDM, ListeChp, _Enr);
    end
    else Result := dmErr_nondef;
    if Result = dmErr_ok then
      Pseq := TdmPseq.create(DM, ListeChp, true)
    else
      Pseq := nil;
  except
    on E: Exception do
    begin
      //result:=100;
      raise(exception.create('Pb dm_pseq dm n° ' + IntToStr(NomDm) + E.Message));
    end;
  end;
end;

function dm_pseqAll(NomDM: noDM; var Pseq: TdmPseq): integer;
var DM: TDataMem;
begin
  try
    if RechercheDM(NomDM, DM)
      then
    begin
      if DM.listCase.count = 0
        then Result := dmErr_nontrouve
      else
      begin
        DM.cursdm := 0;
        DM.rempliBuff(0); // lit la 1ère cas si liste vide
        Result := dmErr_ok;
      end;
    end
    else Result := dmErr_nondef;
    if Result = dmErr_ok then
      Pseq := TdmPseq.create(DM, ListeNil, false)
    else
      Pseq := nil;
  except
    on E: Exception do
    begin
      //result:=100;
      raise(exception.create('Pb dm_pseqall dm n° ' + IntToStr(NomDm) + E.Message));
    end;
  end;
end;

function dm_pseqLect(Pseq: tdmPseq; var Enreg: Tenreg): integer;
var DM: TDataMem;
  OK: boolean;
  C: TLList;
  i: integer;
begin
  try
    Result := dmErr_finenreg;
    if not (assigned(Pseq)) then exit;
    DM := Pseq.dm;
    if Pseq.debut then
      DM.cursdm := Pseq.curseur // on prend le 1er au début
    else
      DM.cursdm := 1 + Pseq.curseur; // on avance le curseur à partir du 2ième appel
    if (DM.listCase.count = 0) or (DM.Cursdm >= DM.listCase.count)
      then
      Result := dmErr_finEnreg
    else
    begin
      OK := True;
      C := DM.ListCase.Items[DM.Cursdm];
      if Pseq.yaFiltre then
        for i := 0 to Pseq.Filtre.Count - 1 do
          if strPas(PChar(C.Items[i + 1])) <> Pseq.Filtre.Strings[i]
            then
          begin
            OK := False;
            Enreg := nil;
            Result := dmErr_finEnreg;
            break;
          end;
      if OK then
      begin
        DM.rempliBuff(DM.CursDm);
        Enreg := DM.buff;
        Result := dmErr_ok;
      end;
    end;
    if (Result = dmErr_ok)
      then
    begin
      Pseq.curseur := DM.cursdm;
      Pseq.debut := false
    end
    else Pseq.free;
  except
    on E: Exception do
    begin
      //result:=100;
      raise(exception.create('Pb dm_pseqlect ' + E.Message));
    end;
  end;
end;

function dm_suivant(NomDm: noDM; var Enreg: Tenreg): integer;
var DM: TDataMem;
begin
  if RechercheDM(NomDM, DM)
    then
  begin
    DM.CursDm := DM.CursDm + 1;
    if (DM.listCase.count = 0) or (DM.Cursdm >= DM.listCase.count)
      then Result := dmErr_finEnreg
    else
    begin
      DM.rempliBuff(DM.CursDm);
      Enreg := DM.buff;
      Result := dmErr_ok;
    end;
  end
  else Result := dmErr_nondef;

end;

procedure TDataMem.addSql(const Sql: string);
var TabEnreg: array of string;
  i: integer;
  requete: TQuery;
  NbChp: integer;
  Val: string;
begin
  requete := opensql(Sql, true);
  requete.First;
  NbChp := requete.FieldCount - 1;
  SetLength(TabEnreg, NbChp + 1);
  while not requete.EOF do
  begin
    for i := 0 to NbChp do
    begin
      Val := requete.Fields[i].AsString;
      if (Val <> '') and (requete.Fields[i].dataType = ftDatetime) then
        if strToDateTime(Val) < 10 then Val := ''; // date en 1899
      TabEnreg[i] := Val;
    end;

    // A cause du problème que 'A_' > 'AR' en Delphi mais pas en sql
    // différence de tri entre Oracle et MS-SQL
    {$IFDEF DATAMEMSANSTIRETBAS}
    add(TabEnreg);
    {$ELSE}
    dm_creaenreg(NumDM, tabenreg);
    {$ENDIF}

    requete.Next;
  end;
  ferme(requete);
end;

function TDataMem.okCle(Lacase: TLList; const ListeChpCle: array of string): integer;
var I, Sens, Max: integer;
  Ch1, Ch2: string;
  Flo1, Flo2: double;
  Dat1, Dat2: TDateTime;
begin
  Sens := 0;
  Max := pluspetit(NbCle, high(ListeChpCle) + 1);
  for I := 1 to Max do
    if TypeChamp[I - 1] = 'A'
      then
    begin
      Ch1 := strPas(PChar(Lacase.items[I]));
      Ch2 := ListeChpCle[I - 1];
      if (UpperCase(Ch1) > UpperCase(Ch2))
        then
      begin
        Sens := -1;
        break;
      end
      else
        if (UpperCase(Ch1) < UpperCase(Ch2))
        then
      begin
        Sens := +1;
        break;
      end
    end
    else
      if TypeChamp[I - 1] = 'N'
      then
    begin
      Flo1 := Valeur(strPas(PChar(Lacase.items[I])));
      Flo2 := Valeur(ListeChpCle[I - 1]);
      if (Flo1 > Flo2)
        then
      begin
        Sens := -1;
        break;
      end
      else
        if (Flo1 < Flo2)
        then
      begin
        Sens := +1;
        break;
      end
    end
    else
      if TypeChamp[I - 1] = 'D'
      then
    begin
      Dat1 := strToDate(strPas(PChar(Lacase.items[I])));
      Dat2 := strToDate(ListeChpCle[I - 1]);
      if (Dat1 > Dat2)
        then
      begin
        Sens := -1;
        break;
      end
      else
        if (Dat1 < Dat2)
        then
      begin
        Sens := +1;
        break;
      end
    end;
  Result := Sens;
end;

// rempli le TLList cledm avec les Pchar sur les valeurs des chp clés
// ls champs clé sont à partir de l'indice 1

procedure TdataMem.genCledm(Lacase: TLList; const TabEnr: array of string);
var Lim, I: integer;
begin
  Lim := nbCle - 1;
  if (high(TabEnr) < Lim) then
    Lim := high(TabEnr);
  for i := 0 to Lim do
    Lacase.items[I + 1] := donneRef(I, TabEnr[I]); // pointe sur la valeur du champ
end;

function Recherche_Code_ListRef_i(Liste: TList; const S: string; deb, fin: integer; var Pos: integer): boolean;
var Sc: string;
  Mil: integer;
begin
  if deb > fin then
  begin
    Pos := deb;
    Result := false;
    exit;
  end;
  Mil := (fin + deb) div 2;
  Sc := strPas(Liste.Items[Mil]);
  if Sc = S then
  begin
    Pos := Mil;
    result := true;
    exit;
  end;
  if Sc < S then
    result := Recherche_Code_ListRef_i(Liste, S, Mil + 1, fin, Pos)
  else
    result := Recherche_Code_ListRef_i(Liste, S, deb, Mil - 1, Pos);
end;

function TdataMem.donneRef(IndCle: integer; const S: string): PChar;
var Lref: Tlist;
  Pos: integer;
  P: Pchar;
  trouve: boolean;
  noRefC: noRefCle;

begin
  NoRefC := listnoRefCle[IndCle];
  if not (assigned(laListRef[NoRefC])) then laListRef[NoRefC] := TList.create;
  Lref := laListRef[NoRefC];
  //  Lref:=ListRef.items[IndCle];
  trouve := Recherche_Code_ListRef_i(Lref, S, 0, Lref.Count - 1, Pos);
  if trouve
    then result := Lref.items[Pos]
  else
  begin
    P := strNew(PChar(S));
    Lref.Insert(Pos, P);
    Result := P;
  end;
end;

procedure TDataMem.rempliBuff(Curs: integer);
begin
  genList(listCase.items[Curs], buff.list);
end;

procedure TDataMem.genList(Lacase: TLList; Enr: TstringList);
var S, S2, Chp: string;
  // Sst: TstringStream;
  I: integer;
begin
  try
    Enr.clear;
    // on rempli les clé
    for I := 1 to nbCle do
      Enr.add(strPas(PChar(Lacase.items[I])));
    {
    if FOkCompress
      then
        begin
          Sst:=TstringStream.create('');
          BZIP.decompressStream(TMemoryStream(Lacase.items[0]),Sst);
          S:=Sst.dataString;
          Sst.free;
        end
      else
    }
    S := strPas(PChar(Lacase.items[0]));
    // et les champs données ensuite
    while extraitChp(S, Chp, S2, #1) do
    begin
      Enr.add(Chp);
      S := S2;
    end;
  except
    on E: Exception do
    begin
      //result:=100;
      raise(exception.create('Pb dm_genlist ' + E.Message));
    end;
  end;

end;

function dm_GetCurseur(nomDM: noDM): integer;
var DM: TdataMem;
begin
  if RechercheDM(NomDM, DM)
    then
  begin
    Result := DM.CursDm;
  end
  else Result := -1;
end;

procedure dm_SetCurseur(nomDM: noDM; Curseur: integer);
var DM: TdataMem;
begin
  if RechercheDM(NomDM, DM)
    then DM.CursDm := Curseur;
end;

function dm_Nb_Enregistrement(nomDM: noDM): integer;
var DM: TDataMem;
begin
  try
    if RechercheDM(NomDM, DM)
      then
      Result := DM.ListCase.Count
    else
      Result := -1;
  except
    on E: Exception do
    begin
      //result:=100;
      raise(exception.create('Pb dm_Nb_Enregistrement : dm n° ' + inttostr(nomDM) + E.Message));
    end;
  end;
end;

{$IFDEF TRACE}

procedure dm_debug(nomDM: noDM);
var Pseq: TdmPseq;
  Enr: Tenreg;
  OK: integer;
  S, Sep: string;
  I, K, curs: integer;
begin
  curs := dm_GetCurseur(nomDM);
  Ok := dm_pseqAll(nomDM, Pseq);
  i := 1;
  debug('--------- Contenu du DM N°' + inttoStr(nomDM) + ' ------------');
  while (Ok = 0) and (dm_pseqLect(Pseq, Enr) = 0) do
  begin
    S := inttoStr(I) + ':';
    inc(i);
    for K := 1 to Enr.nbchamp do
    begin
      if K > 1 then Sep := ' ; ' else Sep := ' ';
      S := S + Sep + Enr.ch(K);
    end;
    debug(S);
  end;
  debug('-------------------------------------------');
  dm_SetCurseur(nomDM, curs);
end;
{$ENDIF}

function dm_SetEnreg(nomDM: noDM; var Enreg: TEnreg; Curseur: integer): integer;
var
  DM: TDataMem;
begin
  try
    if RechercheDM(NomDM, DM)
      then
    begin
      if (Curseur >= DM.ListCase.Count) or (Curseur < 0)
        then Result := dmErr_nontrouve
      else
      begin
        DM.CursDm := Curseur;
        DM.rempliBuff(Curseur);
        Enreg := DM.buff;
        Result := dmErr_ok;
      end;
    end
    else Result := dmErr_nondef;
  except
    on E: Exception do
    begin
      //result:=100;
      raise(exception.create('Pb dm_SetEnreg : dm n° ' + inttostr(nomDM) + E.Message));
    end;
  end;
end;

//supprime l'enregistrement courant

function dm_suppEnregCour(nomDM: noDM): integer;
var DM: TdataMem;
  C: TLList;
  curseur: integer;
begin
  try
    if RechercheDM(NomDM, DM)
      then
    begin
      Curseur := DM.CursDm;
      if (Curseur < DM.ListCase.count) and (Curseur >= 0)
        then
      begin
        C := DM.ListCase[curseur]; // TLList à supprimer
        if DM.FokCompress
          then TMemoryStream(C.items[0]).free // Tstream libéré
        else strDispose(PChar(C.items[0])); // Pchar libéré
        DM.ListCase.Remove(C);
        C.free; // TLList libéré, les Pchar de cette liste sont libérés via Listref
        Result := dmErr_ok;
      end
      else
        Result := dmErr_nontrouve;
    end
    else
      Result := dmErr_nondef;
  except
    on E: Exception do
    begin
      //result:=100;
      raise(exception.create('Pb dm_suppEnregCour : dm n° ' + inttostr(nomDM) + E.Message));
    end;
  end;
end;

//modifie le champ de l'enreg en cours par Chp
//NumChp : numéro du champ à changer

function dm_modifchcour(nomDM: noDM; NumChp: noChamp; const Chp: string): integer;
var DM: TDataMem;
  C: TLList;
  curseur, i, NbList: integer;
  Enreg: TEnreg;
  TabString: array of string;
  Chpi: noChamp;
  S: string;
begin
  try
    if RechercheDM(NomDM, DM)
      then
    begin
      if DM.FOkCompress then
      begin
        raise(exception.create('dm_modifChcour non géré avec compression'));
        Result := dmErr_erreur;
        exit;
      end;
      Curseur := DM.CursDm;
      if (Curseur < DM.ListCase.count) and (Curseur >= 0)
        then
      begin
        // TLList à modifier
        C := DM.ListCase[curseur];
        if ((NumChp - 1) < DM.NbCle)
          // si le champ à changer est une clé
        then
        begin
          raise(exception.create('Pas de modifchcour sur chp en clé'));
          Result := dmErr_erreur;
          exit;
        end
        else
        begin
          // Décompose la TLListe
          DM.rempliBuff(curseur);
          // Change la valeur du champ NumChp
          NbList := DM.Buff.list.Count;
          for i := NbList to NumChp - 1 do
            DM.Buff.list.add('');
          DM.Buff.list[NumChp - 1] := Chp;
          // Regénère la TLListe avec la modif
          Enreg := DM.Buff;
          // Transforme le TEnreg en array of string
          Setlength(TabString, Enreg.nbChamp);
          for chpi := 1 to Enreg.nbChamp() do
          begin
            TabString[chpi - 1] := Enreg.ch(chpi);
          end;
          // Genère un string avec #1
          S := genenreg(DM.NbCle, TabString, #1);
          strdispose(C.items[0]); // libération de l'ancien
          C.items[0] := strnew(PChar(S)); // création du PChar à l'indice 0
          Result := dmErr_OK;
        end
      end
      else
        Result := dmErr_nontrouve;
    end
    else
      Result := dmErr_nondef;
  except
    on E: Exception do
    begin
      //result:=100;
      raise(exception.create('Pb dm_modifChCour : dm n° ' + inttostr(nomDM) + E.Message));
    end;
  end;
end;

//modifie les champs  non en clé sur l'enreg courant

function dm_modifEnregCour(nomDM: noDM; const TabEnr: array of string): integer;
var DM: TDataMem;
  C: TLList;
  curseur: integer;
  S: string;
begin
  try
    if RechercheDM(NomDM, DM)
      then
    begin
      if DM.FOkCompress then
      begin
        raise(exception.create('dm_modifEnregcour non géré avec compression'));
        Result := dmErr_erreur;
        exit;
      end;
      Curseur := DM.CursDm;
      if (Curseur < DM.ListCase.count) and (Curseur >= 0)
        then
      begin
        // TLList à modifier
        C := DM.ListCase[curseur];
        // Genère un string avec #1
        S := genenreg(DM.NbCle, TabEnr, #1);
        strdispose(C.items[0]); // libération de l'ancien
        C.items[0] := strnew(PChar(S)); // création du PChar à l'indice 0
        Result := dmErr_OK;
      end
      else
        Result := dmErr_nontrouve;
    end
    else
      Result := dmErr_nondef;
  except
    on E: Exception do
    begin
      //result:=100;
      raise(exception.create('Pb dm_modifEnregCour : dm n° ' + inttostr(nomDM) + E.Message));
    end;
  end;
end;

//supprime un enregistrement

function dm_suppEnreg(nomDM: noDM; const ListeChpCle: array of string): integer;
var _Enr: TEnreg;
  DM: TDataMem;
  Trouve: integer;
begin
  try
    if RechercheDM(NomDM, DM)
      then
    begin
      Trouve := dm_TrouveEnr(NomDM, ListeChpCle, _Enr);
      if Trouve = dmErr_OK
        then
        // si l'enregistrement existe on supprime l'enr en cours
        // car dm_TrouveEnr place le curseur sur l'enr trouvé
        // mais la fonction SuppEnregCour supprime à Cursdm-1
      begin
        dm_SuppEnregCour(NomDM);
        Result := dmErr_OK;
      end
      else
        Result := Trouve; // nontrouvé ou nondef
    end
    else
      Result := dmErr_nondef;
  except
    on E: Exception do
    begin
      //result:=100;
      raise(exception.create('Pb dm_suppEnreg : dm n° ' + inttostr(nomDM) + E.Message));
    end;
  end;

end;

//ajoute un enregistrement

function dm_creaEnreg(nomDM: noDM; const TabEnr: array of string): integer;
var DM: TDataMem;
  i, Item: integer;
  S: string;
  C: TLList;
begin
  try
    if RechercheDM(NomDM, DM)
      then
    begin
      if DM.FOkCompress
        then
      begin
        showmessage('dm_creaenreg non géré avec compression');
        Result := dmErr_erreur;
        exit;
      end;
      Item := DM.Rech_recurs(TabEnr, 0, DM.ListCase.Count - 1, i);
      if (Item <> -1) and (DM.cleunique) then
      begin
        result := dmErr_keyViol;
        exit;
      end;
      // Création de la TLListe
      C := TLList.create(DM.NbCle + 1);

      // Insertion de la TTliste crée dans ListCase à l'index
      DM.ListCase.insert(i, C);
      // Génération d'une chaine comprenant l'enr avec #1
      S := genenreg(DM.NbCle, TabEnr, #1);
      C.items[0] := strnew(PChar(S)); // création du PChar à l'indice 0
      DM.genCledm(C, TabEnr);
      DM.cursdm := i;
      Result := dmErr_OK;
    end
    else
      result := dmErr_nondef;
  except
    on E: Exception do
    begin
      //result:=100;
      raise(exception.create('Pb dm_creaEnreg : dm n° ' + inttostr(nomDM) + E.Message));
    end;
  end;

end;

procedure initUdatamem;
var i: integer;
begin
  for i := 1 to maxDatamem do
    ListDatamem[i] := nil;
  for i := 0 to maxNbRefCle do
    laListRef[i] := nil;
end;

function dmErrLibelle(Err: integer): string;
begin
  case Err of
    dmErr_dejaExistant: Result := 'enreg déja existant';
    dmErr_nondef: Result := 'DM non défini'; // n° de datamem non créé
    dmErr_keyViol: Result := 'DM violation de clé'; // n° de datamem non créé
  else Result := 'DM erreur';
  end;
end;

initialization
  initUdatamem;

finalization
  destroyAllDm();

end.
