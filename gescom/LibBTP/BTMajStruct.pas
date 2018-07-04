unit BTMajStruct;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
	ComCtrls, StdCtrls, Hctrls, ExtCtrls, Hgauge,UTob,EntGc,HEnt1, UBob, FileCtrl,uHListe,
{$IFDEF EAGLCLIENT}
     UtileAGL,uHttp,uWaini,
{$ELSE}
     DBGrids,
     DBCtrls,
     DB,
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     MajTable,MajHalleyUtil,
     HDB,

{$ENDIF}
		 uBTPVerrouilleDossier,
     UtilPGI,
     
     HMsgBox, TntStdCtrls ;

type

  TfbtMajStructure = class(TForm)
    Titre: THLabel;
    PB: TEnhancedGauge;
    CURVUE: THLabel;
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

  TMajStructBTP = class
  private
    fCodeStatus : integer;
    TheListModifDest,TheListModifRef : TOB;
    DBRef : TDatabase;
    DestDriver: TDBDriver;
    DestODBC: Boolean;
    RefDriver : TDBDriver;
    RefODBC : Boolean;
    IsFirstChange : boolean;
    IsModeMajHalley : boolean;
    XX :  TfbtMajStructure;
    LanceMajVues : boolean;
    VersionBaseRef, VersionBaseDest : string;
    ForceGlobale : boolean;
    //
    procedure TraiteChangementTable (TheListRef : TOB);
    procedure TraiteChangementParamSoc (TheListRef : TOB);
    procedure GetStructureTableFromRef (TheSTructureRef : TOB; NomTable : string);
    procedure TraiteChangementMenu (TheListRef : TOB);
    procedure TraiteModifVue (TheListRef : TOB);
    procedure CreeTable(TheStructureRef: TOB; NomTable: string);
    function ModifieTable(TheListRef,TheStructureRef: TOB; NomTable: string): boolean;
    procedure InitChampsAjoute (NomTable : string);
    function IsATraiter (NomTable : string; Force : boolean) : boolean;
    procedure SetCodeStatus(const Value: integer);
    procedure GetParamSocs (TheListRef,TheParamSocsRef,TheParamSocsDest : TOB);
    function GetRequete(TheListRef: TOB; DBREF: TDatabase): string;
    procedure UpDateDecoupeLigne(SetSQL: string; SetCondition: string='');
    procedure UpDateDecoupeLigneBase(SetSQL: string; SetCondition: string='');
    procedure UpDateDecoupeLigneOuvPlat(SetSQL: string; SetCondition: string='');
    procedure UpDateDecoupeLigneCompl(SetSQL:string; SetCondition: string='');
    procedure UpDateDecoupeLigneFac(SetSQL : string; SetCondition: string='');
    procedure UpDateDecoupeLigneOUV(SetSQL:string; SetCondition: string='');
    procedure UpDateDecoupePiece(SetSQL:string; SetCondition: string='');
    procedure UpDateDecoupePiedBase(SetSQL:string; SetCondition: string='');
    procedure UpDateDecoupePieceRG(SetSQL:string; SetCondition: string='');
    function GetVueRef(NomVue: string): TOB;
    function IsTraiteVue(TheListRef,TOBRef, TOBDest: TOB): boolean;
    function GetVueDestination(NomVue: string): TOB;
    procedure DropVuesPhysiques;
    procedure CreateVuesPhysiques;
    procedure GetMenus(NumeroMenu: string; TheMenusRef,TheMenusDest: TOB);
    procedure GetMenusSpecif (NumeroMenu: string; TheMenusSpecif : TOB);

    procedure DropMenuDest(NumeroMenu: string);
    procedure beforeTraitement;
    procedure AfterTraitement;
    procedure ControleTableEchange;
    procedure CreateSqlFunctions;
    procedure CreateFunctionDecoupage;
    procedure DropFunction(NomF: String);
    procedure Updateparamsocs;
    procedure GetAllParamSocs(TheParamSocsRef, TheParamSocsDest: TOB);
    procedure AddmenusSpecif(TheMenusRef, TheMenusSpecif: TOB);
    procedure TraiteModifStructureListe(TheListRef: TOB);
    procedure ConstitueTableAdressesBTP;
    procedure ConstitueTableContactBTP;
    procedure ConstitueTableFonctionBTP;
    procedure ConstitueTableRessourceBTP;
    procedure ConstitueTableSoucheBTP;
    procedure ConstitueTableTiersBTP;
    procedure CreateRealTable(NomTable: string);
    procedure InitChampsUniqueBAST;

  public
    property CodeStatus : integer read fCodeStatus write SetCodeStatus;
    constructor create;
    destructor destroy; override;
    procedure lance;
    function IsTablesATraiter : boolean;

  end;


const TEMPTABLE = 'TEMPOR';


function MajStructure (IsModeMajHalley : boolean=false): integer;
procedure TraitementsAfterImportBOB (Version : string);
procedure TraitementsBeforeImportBOB (VersionBase : string);
function ISColumnExists(Champs : string; Table : string) : boolean;

implementation
uses ChangeVersions,Paramsoc,uPGFnWorkRh,PGInitRemDSN,  ed_tools
;
{$R *.DFM}

function ISColumnExists(Champs : string; Table : string) : boolean;
begin
  result := ExisteSQL ('SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE COLUMN_NAME="'+Champs+'" and TABLE_NAME="'+Table+'"');
end;


function ExecuteSQLContOnExcept(const Sql: WideString): Integer;
begin
  result := -1;

  try
    result := ExecuteSQL(sql);
  except
    on e: exception do
    begin
      if  (v_pgi.SAV) then PgiError(E.message);
    end;
  end;

end;



procedure TraitementsBeforeImportBOB (VersionBase : string);
begin
  if VersionBase < '998.ZR' then // edition 15  paie
  begin
		ExecuteSQL('DELETE FROM MOTIFSORTIEPAY WHERE PMS_CODE IN ("70", "71")');
		ExecuteSql('DELETE FROM YEDMFILES WHERE YEF_ID IN (SELECT YEM_ID FROM YEDMEXTMETIER WHERE YEM_PREDEFINI = "CEG")');
		ExecuteSQL('DELETE FROM YEDMEXTMETIER WHERE YEM_PREDEFINI = "CEG"');
		ExecuteSql('DELETE FROM COMMUN  WHERE CO_TYPE = "YEB" AND CO_LIBRE="PP8"');
    ExecuteSQL('DELETE FROM COMMUN  WHERE CO_TYPE = "YBU" AND CO_CODE="PP8"');
    ExecuteSQL('DELETE FROM PARAMSALARIE WHERE PPP_PREDEFINI = "CEG"');
  end;
  if VersionBase < '998.ZZ5' then
  begin
    ExecuteSQL('DELETE FROM INSTITUTIONPAYE WHERE PIP_PREDEFINI = "CEG"');
    ExecuteSQL('DELETE FROM DSNAFFECTATION WHERE PNF_PREDEFINI = "CEG"');
    ExecuteSQL('DELETE FROM DSNDONNEES');
    ExecuteSQL('DELETE FROM DSNDECLARATIONS');
    ExecuteSQL('DELETE FROM DSNLEXIQUE');
    ExecuteSQL('DELETE FROM DSNFONCTION WHERE PFN_PREDEFINI = "CEG"');
    ExecuteSQL('DELETE FROM DSNFONCTIONLG WHERE PFL_PREDEFINI = "CEG"');
  end;
end;

procedure TraitementsAfterImportBOB (Version : string);
var QQ : TQuery;
		res,SQL : string;
begin
  if Version < '998.ZR' then // edition 15
  begin
  	QQ := OpenSQL('SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_PGTYPECONTRAT"',True,1,'',true);
    if not QQ.eof then
    begin
      res := QQ.fields[0].AsString;
      if (pos(Res,'CCD;CCO;CCS')>0) then
      begin
      	ExecuteSQL ('UPDATE PARAMSOC SET SOC_DATA="01" WHERE SOC_NOM="SO_PGMOTIFRECOURSCDD"');
      end else
      begin
      	ExecuteSQL ('UPDATE PARAMSOC SET SOC_DATA="" WHERE SOC_NOM="SO_PGMOTIFRECOURSCDD"');
      end;
    end;
    Ferme(QQ);
    //
    SQL := 'SELECT 1 FROM REMUNERATION WHERE PRM_PREDEFINI <> "CEG" AND PRM_NATUREPRIME <> "" ' +
    			 'UNION ' +
    			 'SELECT 2 FROM REMUNERATION WHERE PRM_PREDEFINI <> "CEG" AND PRM_AUTREELEMENT <> ""';
    if not ExisteSQL(SQL) then
    begin
      InitRemDSN;
    end;
  end;
  if Version < '998.ZZ5' then
  begin
    ExecuteSQL ('delete from menu where mn_1 = "41" and mn_2 = "13" and mn_tag like "%419%"');
    ExecuteSQL ('delete from menu where mn_1 = "48" and mn_2 = "3"');
    ExecuteSQL ('update commun set co_libre = " " where co_type="PZN" and co_libre is null');
  end;
end;

function MajStructure (IsModeMajHalley : boolean=false): integer;
var TheMajStruct : TmajStructBTP;
begin
  //
	if not IsModeMajHalley then
  begin
    if ISVerrouille Then
    begin
      PGIInfo ('Un traitement de mise à jour est en cours..#13#10 Merci de vous reconnecter ultérieurement.');
      result := -1 ;
      exit;
    end;
  end;
  //
  TheMajStruct := TMajStructBTP.create;
  TRY
  	TheMajStruct.IsModeMajHalley := IsModeMajHalley;
    if TheMajStruct.IsTablesATraiter then
    begin
      //
      if not IsModeMajHalley then
      begin
        if PGIAsk  ('Le produit nécessite une mise à jour de structure.#13#10 Avez-vous fait la sauvegarde ?')=MrYes then
        begin
          if PgiAsk ('Confirmez-vous le traitement ?')=Mryes then
          begin
            TheMajStruct.Lance;
            If TheMajStruct.CodeStatus = 1 then
              PGIInfo ('Traitement de mise à jour terminé avec succès.');
          end else TheMajStruct.CodeStatus := -1;
        end else TheMajStruct.CodeStatus := -1;
      end else
      begin
        TheMajStruct.Lance;
      end;
    end;
  FINALLY
    result := TheMajStruct.CodeStatus;
    TheMajStruct.ControleTableEchange;
    TheMajStruct.free;
		if not IsModeMajHalley then
    begin

    	Deverrouille;
    end;
  END;
end;


{ TMajStructBTP }

procedure TMajStructBTP.GetStructureTableFromRef(TheSTructureRef: TOB;NomTable: string);
var QQ : TQuery;
begin
  QQ := OpenSQLDb (DBREF,'SELECT * FROM DETABLES WHERE DT_NOMTABLE="'+NomTable+'"',true);
  if QQ.eof then BEGIN fCodeStatus := 120; Exit; END;
  TheSTructureRef.SelectDB ('',QQ);
  ferme (QQ);
  QQ := OpenSQLDb (DBREF,'SELECT * FROM DECHAMPS WHERE DH_PREFIXE="'+TheSTructureRef.getValue('DT_PREFIXE')+'"',true);
  TheSTructureRef.LoadDetailDB ('DECHAMPS','','',QQ,false);
  ferme (QQ);
end;

function TMajStructBTP.IsTablesATraiter: boolean;
var QQ : TQuery;
    TheModifDest,TheModifRef : TOB;
begin
  if IsModeMajHalley then BEGIN result := true; exit; END;
  result := false;
  if TableExiste ('BTMAJSTRUCTURES') then
  begin
    QQ := OpenSql ('SELECT * FROM BTMAJSTRUCTURES',true);
    TheListModifDest.LoadDetailDB ('BTMAJSTRUCTURES','','',QQ,false);
    ferme (QQ);
  end;
  if TableExisteDB  ('BTMAJSTRUCTURES',DBREF) then
  begin
    QQ := OpenSQLDb (DBREF,'SELECT * FROM BTMAJSTRUCTURES',true);
    TheListModifRef.LoadDetailDB('XXX','','',QQ,false);
    ferme (QQ);
  end;

  if TheListModifRef.detail.count > 0 then
  begin
    VersionBaseRef := '';
    VersionBaseDest := '';
    //
    TheModifRef := TheListModifRef.findFirst(['BTV_TYPEELT'],['VERSION'],True);
    versionBaseRef := Trim(TheModifRef.getValue('BTV_VERSIONBASEB'));
    TheModifDest := TheListModifDest.findFirst(['BTV_TYPEELT'],['VERSION'],True);
    if TheModifDest <> nil then
    begin
      VersionBaseDest := Trim(TheModifDest.GetValue('BTV_VERSIONBASEB'));
    end;
    if (VersionBaseRef <= VersionBaseDest) then exit;
  end else exit;
  result := True;
end;

constructor TMajStructBTP.create;
begin
  IsFirstChange := true;
  fCodeStatus := 0;
  TheListModifDest  := TOB.create ('LES TABLES MODIFIEES PAR BATIMENT',nil,-1);
  TheListModifRef := TOB.create ('LES TABLES A MODIFIER',nil,-1);
  // Sauvegarde du contexte
  DestDriver := V_PGI.Driver; DestODBC := V_PGI.ODBC; V_PGI.StopCourrier := TRUE;
  //
  DBRef := TDatabase.Create(Application);
  DBRef.Name := 'DBREF';
  // Overture de la base REF
  ConnecteDB('Reference', DBRef, 'DBREF');
  RefDriver := V_PGI.Driver; RefODBC := V_PGI.ODBC;
  //
  XX := TfbtMajStructure.Create(Application);  // partie IHM
  XX.Visible := false;
end;

destructor TMajStructBTP.destroy;
begin
  XX.free;
  DBRef.Connected := False;
  DBRef.Free;
  //
  TheListModifDest.free;
  TheListModifRef.free;
  //
  V_PGI.Driver := DestDriver; V_PGI.ODBC := DestODBC; V_PGI.StopCourrier := FALSE;

  inherited;
end;

procedure TMajStructBTP.lance;
var Indice : integer;
    TheListRef : TOB;
    QQ : TQUery;
    VersionElt : string;
begin
//  if not IsModeMajHalley then
  begin
    if VersionBaseDest = '998.0' then
    begin
      ExecuteSQLContOnExcept('delete from ymybobs where yb_bobname like "BAT30999%"');
      ForceGlobale := True;
    end;
  	XX.Visible := true;
  	V_Pgi.EnableDeShare := False;
    //
    beforeTraitement;
    if fCodeStatus <> 0 then Exit;
    //
    LanceMajVues := False;

    for Indice := 0 to TheListModifRef.detail.count -1 do
    begin
      TheListref := TheListModifRef.detail[Indice];
      VersionElt := Trim(TheListRef.GetString('BTV_VERSIONBASEB'));
      // controle si l'element en cours n'est pas déja traité
      if (VersionElt <= VersionBaseDest ) then Continue;
      // Traitement des éléments jusqu'a la version de la base ref
      if ( VersionElt <= VersionBaseRef ) then
      begin
        if TheListRef.GetValue('BTV_TYPEELT')='VERSION' then
        begin
          continue;
        end else if TheListRef.GetValue('BTV_TYPEELT')='TABLE' then
        begin
          TraiteChangementTable (TheListRef);
        end else if TheListRef.GetValue('BTV_TYPEELT')='VUE' then
        begin
          TraiteModifVue (TheListRef);
        end else if TheListRef.GetValue('BTV_TYPEELT')='PARAMSOC' then
        begin
          TraiteChangementParamSoc (TheListRef);
        end else if TheListRef.GetValue('BTV_TYPEELT')='MENU' then
        begin
          TraiteChangementMenu (TheListRef);
        end else if TheListRef.GetValue('BTV_TYPEELT')='LISTE' then
        begin
          TraiteModifStructureListe (TheListRef);
        end;
        if fCodeStatus > 1 then break;
      end;
    end;
    if (fCodeStatus <= 1) then
    BEGIN
      AfterTraitement;

      V_Pgi.EnableDeShare := true;
      if LanceMajVues then
      begin
        DropVuesPhysiques ;
        CreateVuesPhysiques;
      end;
    END;
    XX.visible := false;
(*
  end else
  begin
    QQ := OpenSQLDb (DBREF,'SELECT * FROM BTMAJSTRUCTURES',true);
    TheListModifRef.LoadDetailDB ('BTMAJSTRUCTURES','','',QQ,false);
    ferme (QQ);

  	ExecuteSql ('DELETE FROM BTMAJSTRUCTURES');
    TheListModifDest.dupliquer (TheListModifRef,true,true);
    TheListModifDest.InsertDB(nil);

    Updateparamsocs;
*)
  end;
  if IsModeMajHalley then Updateparamsocs;

end;

procedure TMajStructBTP.TraiteChangementParamSoc(TheListRef: TOB);
var NomParamSoc : string;
    TheParamSocsRef,TheParamSocsDest : TOB;
    TheParamSocRef,TheParamSocDest : TOB;
    indice : integer;
    QQ : TQuery;
begin
  TheParamSocsref := TOB.create ('LES PARAMS SOC',nil,-1);
  TheParamSocsDest := TOB.create ('LES PARAMS SOCD',nil,-1);

  NomparamSoc := TheLIstRef.GetValue ('BTV_NOMELT');
	XX.Titre.Caption := 'Mise à jour de Paramètres Sociétés';
  XX.Titre.visible := true;
  XX.Titre.Refresh;
  XX.PB.Visible := true;
  XX.Refresh;
  GetParamSocs (TheListRef,TheParamSocsRef,TheParamSocsDest);
  for indice := 0 to TheParamSocsref.detail.count -1 do
  begin
    TheParamSocRef :=TheParamSocsref.detail[Indice];
    TheParamSocDest := TheParamSocsDest.findFirst(['SOC_NOM'],[TheParamSocRef.GetValue('SOC_NOM')],true);
    if TheParamSocDest = nil then
    begin
      TheParamSocDest := TOB.Create('PARAMSOC',TheParamSocsDest,-1);
      QQ := OpenSql ('SELECT * FROM PARAMSOC WHERE SOC_NOM="'+TheParamSocRef.GetValue('SOC_NOM')+'"',true);
      if QQ.eof then
      begin
        // Nouveau paramsoc
        TheParamSOcDest.Dupliquer (TheParamSocRef,false,true);
        TheParamSocDest.SetAllModifie (true);
      end else
      begin
        // Cas particulier du changement de SOC_TREE
        TheParamSocDest.selectDB('',QQ);
        TheParamSocDest.putValue('SOC_DESIGN',TheParamSocRef.GetValue('SOC_DESIGN'));
        TheParamSocDest.putValue('SOC_TREE',TheParamSocRef.GetValue('SOC_TREE'));
        if TheListRef.GetValue('BTV_WITHDATA')='X' then
        begin
          TheParamSocDest.putValue('SOC_DATA',TheParamSocRef.GetValue('SOC_DATA'));
        end;
        TheParamSocDest.SetAllModifie (true);
      end;
      ferme (QQ);
    end else
    begin
      TheParamSocDest.putValue('SOC_DESIGN',TheParamSocRef.GetValue('SOC_DESIGN'));
      TheParamSocDest.putValue('SOC_TREE',TheParamSocRef.GetValue('SOC_TREE'));
      if TheListRef.GetValue('BTV_WITHDATA')='X' then
      begin
        TheParamSocDest.putValue('SOC_DATA',TheParamSocRef.GetValue('SOC_DATA'));
      end;
      TheParamSocDest.SetAllModifie (true);
    end;
  end;
  //
  if TheParamSocsref.detail.count <> 0 then
  begin
    TheParamSocsDest.InsertOrUpdateDB (true);
  end;
  //
  TheParamSocsref.free;
  TheParamSocsDest.free;
  //
end;

procedure TMajStructBTP.TraiteChangementTable(TheListRef: TOB);
var NomTable : string;
    TheStructureRef : TOB;
begin

  NomTable := TheLIstRef.GetValue ('BTV_NOMELT');
	XX.Titre.Caption := 'Mise à jour de '+NomTable;
  XX.Titre.visible := true;
  XX.Titre.Refresh;
  XX.PB.Visible := true;
  XX.Refresh;
  //
  TheStructureRef := TOB.Create ('DETABLES',nil,-1);
  //
  if TableExisteDB (NomTable,DBREF) then
  begin
    //
    //
    fCodeStatus := 0;
    GetStructureTableFromRef (TheSTructureRef,NomTable);
    if fCodeStatus <> 0 then exit;
    if not TableExiste (NomTable) then
    begin
      XX.PB.MinValue := 0;
      XX.PB.MaxValue  := 3;
      CreeTable (TheStructureRef,NomTable);
      if fCodeStatus <> 0 then exit;
      if TheListRef.GetValue('BTV_WITHDATA')='X' then
      begin
        DBCOPYTABLEBM(DBREF, DBSOC, NomTable, NomTable);
        XX.PB.Progress := XX.PB.Progress + 1;
      end;
      LanceMajVues := True;
    end else
    begin
      XX.PB.MinValue := 0;
      XX.PB.MaxValue  := 8;
      if ModifieTable (TheListRef,TheStructureRef,NomTable) then
      begin
        InitChampsAjoute (NomTable);
        LanceMajVues := True;
      end;
      XX.PB.Progress := XX.PB.Progress + 1;
      if NomTable = 'BTMAJSTRUCTURES' then
      begin
        ExecuteSql ('DELETE FROM BTMAJSTRUCTURES');
        DBCOPYTABLEBM(DBREF, DBSOC, NomTable, NomTable);
      end;
    end;
  end;
  //
  fcodeStatus := 1;
  TheStructureRef.free;
end;

procedure TMajStructBTP.TraiteModifStructureListe(TheListRef: TOB);

  procedure ModifieListe (TOBListe : TOB; QQ : TQuery);
  var TS,TD: TStringList;
  begin
    TS:= TStringList.Create;
    TD := TStringList.Create;
    try
      TS.SetText(PChar(QQ.FindField('LI_DATA').AsString));
      TD.SetText(PChar(TOBListe.getString('LI_DATA')));
      if TS.Count > 0 then
      begin
        TD.Strings[0] := TS.Strings[0]; // le nom de la table ou de la vue
        TD.Strings[2] := TS.Strings[2]; // les jointures
        TOBListe.setString('LI_DATA',TD.GetText); 
      end;
    finally
      TS.Free;
      TD.free;
    end;
  end;

var NomListe : string;
    QQ,QQD : TQuery;
    TOBLISTES,TL : TOB;
    II : integer;
begin
  TOBLISTES:=TOB.Create ('LES LISTES',nil,-1);
  NomListe := TheLIstRef.GetValue ('BTV_NOMELT');
	XX.Titre.Caption := 'Mise à jour de la présentation '+NomListe;
  XX.Titre.visible := true;
  XX.Titre.Refresh;
  XX.PB.Visible := true;
  XX.Refresh;
  //
  QQ := OpenSQLDb(DBREF,'SELECT * FROM LISTE WHERE LI_LISTE="'+NomListe+'" AND LI_LANGUE="FRA" AND LI_UTILISATEUR="---"',true,1,'',true);
  if not QQ.eof then
  begin
    //
    QQD := OpenSQL('SELECT * FROM LISTE WHERE LI_LISTE="'+NomListe+'"',true,1,'',true); // on prends la totalité des listes du client
    if not QQ.eof then
    begin
      TOBLISTES.LoadDetailDB('LISTE','','',QQD,false);
    end;
    ferme (QQD);
    if TOBLISTES.detail.count > 0 then
    begin
      XX.PB.MinValue := 1;
      XX.PB.MaxValue  := TOBLISTES.detail.count;
      for II := 0 To TOBLISTES.detail.count -1 do
      begin
        ModifieListe(TOBLISTES.detail[II],QQ);
        TOBLISTES.detail[II].UpdateDB; 
        XX.PB.Progress := XX.PB.Progress + 1;
      end;
    end else
    begin
      TL := TOB.create ('LISTE',TOBLISTES,-1);   // creation dela liste dans la base de donnée dest
      TL.SelectDB('',QQ);
      TL.SetAllModifie(true);
      TL.InsertDB(nil);
    end;
    //
    fCodeStatus := 0;
  end;
  ferme (QQ);
  TOBLISTES.free;
  //
  fcodeStatus := 1;
end;



procedure TMajStructBTP.CreeTable (TheStructureRef : TOB;NomTable : string);
begin
  DbCreateTable(DBSOC, TheStructureRef, DestDriver, false);
  XX.PB.Progress := XX.PB.Progress + 1;
//  DBCREATEINDEX(DBSOC, TheStructureRef, DestDriver);
  XX.PB.Progress := XX.PB.Progress + 1;
  // Mise à jour de DETABLES ET DECHAMPS
  TheStructureRef.SetAllModifie (true);
  TheSTructureRef.InsertDB (nil);
end;

function TMajStructBTP.ModifieTable (TheListRef,TheStructureRef : TOB;NomTable : string) : boolean;
var Force : Boolean;
begin
	result := false;
  Force := (TheListRef.GetString('BTV_FORCE') = 'X') or (ForceGlobale);
  if not IsATraiter (NomTable,Force) then exit;
  result := true;
  if TableExiste (TEMPTABLE) then DBDeleteTable(DBSOC,DestDriver, TEMPTABLE, TRUE);
  XX.PB.Progress := XX.PB.Progress + 1;
  // Phase 2 : Création de la table temporaire avec la nouvelle structure
  DbCreateTable(DBSOC, TheStructureRef, DestDriver, TRUE);
  XX.PB.Progress := XX.PB.Progress + 1;

  // Phase 3 : Copie de l'ancienne table dans la nouvelle table
  DBCOPYTABLEBM(DBSOC, DBSOC, NomTable, TEMPTABLE);
  XX.PB.Progress := XX.PB.Progress + 1;

  // Suppression de DETABLES et DECHAMPS
  ExecuteSql('DELETE FROM DECHAMPS WHERE DH_PREFIXE="' +TheStructureRef.GetValue('DT_PREFIXE') + '"');
  ExecuteSql('DELETE FROM DETABLES WHERE DT_NOMTABLE="' +TheStructureRef.GetValue('DT_NOMTABLE') + '"');
  XX.PB.Progress := XX.PB.Progress + 1;

  // Phase 4 : Suppression de l'ancienne table
  DBDeleteTable(DBSOC, DestDriver, NomTable, TRUE);
  XX.PB.Progress := XX.PB.Progress + 1;

  // Phase 5 : Renommage de la table temporaire sous le nouveau nom
  DBRenameTable(DBSOC, DestDriver, TEMPTABLE, NomTable);
  XX.PB.Progress := XX.PB.Progress + 1;

  // Phase 6 : Création des index
  //  ConnecteTRUEFALSE(DBSOC);
  DBCREATEINDEX(DBSOC, TheStructureRef, DestDriver);
  XX.PB.Progress := XX.PB.Progress + 1;
  // Mise à jour de DETABLES ET DECHAMPS
  TheStructureRef.SetAllModifie (true);
  TheSTructureRef.InsertDB (nil);
end;

procedure TMajStructBTP.InitChampsAjoute(NomTable: string);
var sql : string;
begin                    
  // ERP CEGID
  //
  if NomTable ='BTRESSOURCE' then
  begin
    if VersionBaseDest < '998.ZZZY' then
    begin
      ExecuteSQL('UPDATE BTRESSOURCE SET BRS_GERECHANTIER="X" WHERE BRS_GERECHANTIER IS NULL');
    end;
  end else if NomTable ='UTILISAT' then
  begin
    if VersionBaseDest < '998.ZZZZ' then
    begin
      ExecuteSQL('UPDATE UTILISAT SET US_RGPDOK="-" WHERE US_RGPDOK IS NULL');
    end;
  end else if NomTable ='PHASESCHANTIER' then
  begin
    if VersionBaseDest < '998.ZZZY' then
    begin
      ExecuteSQL('UPDATE PHASESCHANTIER SET BPC_VISIBLE="X",BPC_SAISISSABLE="X" WHERE BPC_VISIBLE IS NULL');
    end;
  end else if NomTable = 'LIGREAANAL' then
  begin
    if VersionBaseDest < '998.ZZZP' then
    begin
      ExecuteSQL('UPDATE LIGREAANAL SET BLR_QTE=0,BLR_DPA=0 WHERE BLR_QTE IS NULL');
    end;
  end else if NomTable = 'BTFACTST' then
  begin
    if VersionBaseDest < '998.ZZZP' then
    begin
      ExecuteSQL('UPDATE BTFACTST SET BM3_ENVOIMAIL="-" WHERE BM3_ENVOIMAIL IS NULL');
    end;
    if VersionBaseDest < '998.ZZZW' then
    begin
      InitChampsUniqueBAST;
    end;
  end else if NomTable = 'TXCPTTVA' then
  begin
    if VersionBaseDest < '998.ZZZL' then
    begin
      ExecuteSQL('UPDATE TXCPTTVA SET TV_COLLECTIF="" WHERE TV_COLLECTIF IS NULL');
    end;
  end else if NomTable = 'CONDITIONNEMENT' then
  begin
    if VersionBaseDest < '998.ZZZK' then
    begin
      ExecuteSQL('UPDATE CONDITIONNEMENT SET GCO_TIERS="" WHERE GCO_TIERS IS NULL');
    end;
  end else if NomTable = 'PIEDPORT' then
  begin
    if VersionBaseDest < '998.ZZS' then
    begin
      ExecuteSQL('UPDATE PIEDPORT SET GPT_COLLECTIF="", GPT_RETENUEDIVERSE="-" WHERE GPT_COLLECTIF IS NULL');
    end;
  end else if NomTable = 'BAFFAIREATT' then
  begin
    if VersionBaseDest < '998.ZZY' then
    begin
      ExecuteSQL('UPDATE BAFFAIREATT SET BAA_TRAITE="-" WHERE BAA_TRAITE IS NULL');
    end;
  end else if NomTable = 'PORT' then
  begin
    if VersionBaseDest < '998.ZZS' then
    begin
      ExecuteSQL('UPDATE PORT SET GPO_COLLECTIF="", GPO_RETENUEDIVERSE="-" WHERE GPO_COLLECTIF IS NULL');
    end;
  end else if NomTable = 'BTEVENTCHA' then
  begin
    if VersionBaseDest < '998.ZZQ' then
    begin
      ExecuteSQL('UPDATE BTEVENTCHA SET BEC_NUMPHASE=0, BEC_LIBPHASE="" WHERE BEC_NUMPHASE IS NULL');
    end;
  end else if NomTable = 'BTEVENEMENTPLA' then
  begin
    if VersionBaseDest < '998.ZZQ' then
    begin
      ExecuteSQL('UPDATE BTEVENEMENTPLA SET BPL_NUMPHASE=0, BPL_LIBPHASE="" WHERE BPL_NUMPHASE IS NULL');
    end;
  end else if NomTable = 'BTEVENPLAN' then
  begin
    if VersionBaseDest < '998.ZZQ' then
    begin
      ExecuteSQL('UPDATE BTEVENPLAN SET BEP_REFEQUIPE="" WHERE BEP_REFEQUIPE IS NULL');
    end;
  end else if NomTable = 'FONCTION' then
  begin
    if VersionBaseDest < '998.ZZQ' then
    begin
      if ISColumnExists('AFO_BCOLOR','FONCTION') then
      begin
      ExecuteSQL('UPDATE FONCTION SET AFO_GEREPLANNING="X",AFO_BCOLOR=16777215,AFO_COLOR=0 WHERE AFO_BCOLOR IS NULL');
    end;
    end;
  end else if NomTable = 'BLIGNEMETRE' then
  begin
    if VersionBaseDest < '998.ZZO' then
    begin
      ExecuteSQL('UPDATE BLIGNEMETRE SET BLM_UNIQUEBLO=0 WHERE BLM_UNIQUEBLO IS NULL');
    end;
  end else if NomTable = 'BSITUATIONS' then
  begin
    if VersionBaseDest < '998.ZW' then
    begin
      ExecuteSQL('UPDATE BSITUATIONS SET BST_VIVANTE="X" WHERE BST_VIVANTE IS NULL');
    end;
    if VersionBaseDest < '998.ZZ7' then
    begin
      ExecuteSQL('UPDATE BSITUATIONS SET BST_INDICESIT=0 WHERE BST_INDICESIT IS NULL');
    end;
  end else if NomTable = 'BTETAT' then
  begin
    if VersionBaseDest < '998.ZZF' then
    begin
      SQL := 'UPDATE BTETAT SET BTA_TYPEACTION="INT",BTA_GESTIONCONSO="-",BTA_AFFAIRE="",BTA_TIERS="",'+
             'BTA_RESSOURCE="",BTA_VALORISE="-",BTA_AFFECTCHANTIER="-",BTA_AFFECTIERS="-",BTA_AFFECTRESS="-" '+
             'WHERE BTA_TYPEACTION IS NULL';
      ExecuteSQL(SQL);
    end;
  end else if NomTable = 'PAIEPARIM' then
  begin
  end else if NomTable = 'HISTOBULLETIN' then
  begin
    if VersionBaseDest < '998.ZR' then
    begin
    	ExecuteSQL ('UPDATE HISTOBULLETIN SET PHB_DATEDEBRUB="'+UsDateTime (Idate1900)+'", '+
      						'PHB_DATEFINRUB="'+UsDateTime (Idate1900)+'" WHERE PHB_DATEDEBRUB IS NULL');
    end;
  end else if NomTable = 'PGHISTODETAIL' then
  begin
    if VersionBaseDest < '998.ZR' then
    begin
    	ExecuteSQL ('UPDATE PGHISTODETAIL SET PHD_CONTRATTRAV=0 WHERE PHD_CONTRATTRAV IS NULL');
    end;
  end else if NomTable = 'PIECEINTERV' then
  begin
    if VersionBaseDest < '998.ZE' then
    begin
			ExecuteSQL ('UPDATE PIECEINTERV SET BPI_DATECONTRAT=BPI_DATECREATION WHERE BPI_DATECONTRAT IS NULL');
    end;
    if VersionBaseDest < '998.ZZV' then
    begin
			ExecuteSQL ('UPDATE PIECEINTERV SET BPI_FAMILLETAXE="TN",BPI_AUTOLIQUID="-" WHERE BPI_FAMILLETAXE IS NULL');
    end;
    if VersionBaseDest < '998.ZZZM' then
    begin
			ExecuteSQL ('UPDATE PIECEINTERV SET BPI_CODEMARCHE="" WHERE BPI_CODEMARCHE IS NULL');
    end;
	end else if NomTable = 'NATUREPREST' then
  begin
    if VersionBaseDest < '998.Z4' then
    begin
  		ExecuteSQL ('UPDATE NATUREPREST SET BNP_SECTION="" WHERE BNP_SECTION IS NULL');
    end;
  end else if NomTable = 'DEPOTS' then
  begin
    if VersionBaseDest < '998.Z4' then
    begin
  		ExecuteSQL ('UPDATE DEPOTS SET GDE_SECTION="" WHERE GDE_SECTION IS NULL');
    end;
  end else if NomTable = 'ARTICLEDEMPRIX' then
  begin
    if VersionBaseDest < '998.W' then
    begin
  		ExecuteSQL ('UPDATE ARTICLEDEMPRIX SET BDP_DPA=0,BDP_PUHTDEV=0,BDP_SELECTIONNE="-" WHERE BDP_DPA IS NULL');
    end;
  end else if NomTable = 'RESSOURCE' then
  begin
    if VersionBaseDest < '998.Z7' then
    begin
      if ISColumnExists('ARS_SECTION','RESSOURCE') then
      begin
  		ExecuteSQL ('UPDATE RESSOURCE SET ARS_SECTION="" WHERE ARS_SECTION IS NULL');
    end;
    end;
  end else if NomTable = 'NATUREPREST' then
  begin
    if VersionBaseDest < '998.Z7' then
    begin
  		ExecuteSQL ('UPDATE NATUREPREST SET BNP_SECTION="" WHERE BNP_SECTION IS NULL');
    end;
  end else if NomTable = 'BTDOMAINEACT' then
  begin
    if VersionBaseDest < '998.Z7' then
    begin
  		ExecuteSQL ('UPDATE BTDOMAINEACT SET BTD_COEFFG_APPEL=BTD_COEFFG,BTD_COEFMARG_APPEL=BTD_COEFMARG WHERE BTD_COEFFG_APPEL IS NULL');
    end;
  end else if NomTable = 'DEPOT' then
  begin
    if VersionBaseDest < '998.Z7' then
    begin
  		ExecuteSQL ('UPDATE DEPOT SET GDE_SECTION="" WHERE GDE_SECTION IS NULL');
    end;
  end else if NomTable = 'DETAILDEMPRIX' then
  begin
    if VersionBaseDest < '998.W' then
    begin
  		ExecuteSQL ('UPDATE DETAILDEMPRIX SET BD0_DPA=0,BD0_PUHTDEV=0 WHERE BD0_DPA IS NULL');
    end;
  end else if NomTable = 'ACOMPTES' then
  begin
    if VersionBaseDest < '998.W' then
    begin
  		ExecuteSQL ('UPDATE ACOMPTES SET GAC_DATEECHEANCE=GAC_DATEECR WHERE GAC_DATEECHEANCE IS NULL');
    end;
    if VersionBaseDest < '998.X' then
    begin
  		ExecuteSQL ('UPDATE ACOMPTES SET GAC_FOURNISSEUR="" WHERE GAC_FOURNISSEUR IS NULL');
    end;
  end else if NomTable = 'PGEXCEPTIONS' then
  begin
    if VersionBaseDest < '998.W' then
    begin
			ExecuteSQL ('UPDATE PGEXCEPTIONS SET PEN_FERME="-" WHERE PEN_FERME IS NULL');
    end;
    if VersionBaseDest < '998.ZR' then
    begin
    	ExecuteSQL ('UPDATE PGEXCEPTIONS SET PEN_ISRUBPERIOD="-", PEN_PERIODERUB="100", PEN_NATUREPRIME="", '+
      						'PEN_APPLICNATURE="'+UsDateTime (Idate1900)+'", '+
                  'PEN_MODIFNATURE="'+UsDateTime (Idate1900)+'" WHERE PEN_ISRUBPERIOD IS NULL');
    end;
  end else if NomTable = 'ETABCOMPL' then
  begin
    if VersionBaseDest < '998.Q' then
    begin
			ExecuteSQL('UPDATE ETABCOMPL SET ETB_NUMCAISSECP="" WHERE ETB_NUMCAISSECP IS NULL');
    end;
    if VersionBaseDest < '998.ZR' then
    begin
      ExecuteSQL ('UPDATE ETABCOMPL SET ETB_DSNPOINTDEPOT="GEN", ETB_REMUNEXPATRIE="" WHERE ETB_DSNPOINTDEPOT IS NULL');
      ExecuteSQL ('UPDATE ETABCOMPL SET ETB_DSNPOINTDEPOT="MSA" WHERE ETB_MSAACTIVITE <> "" OR ETB_MSASECTEUR  <> "" OR ETB_MSAUNITEGES <> ""');
    end;
  end else if NomTable = 'CONTRATPREV' then
  begin
    if VersionBaseDest < '998.ZR' then
    begin
    ExecuteSQL ('UPDATE CONTRATPREV SET POP_PREVOYANCEOPT="", POP_PREVOYANCEPOP="" WHERE POP_PREVOYANCEOPT IS NULL');
    end;
  end else if NomTable = 'CONTRATTRAVAIL' then
  begin
    if VersionBaseDest < '998.Q' then
    begin
			ExecuteSQL('UPDATE CONTRATTRAVAIL SET PCI_DATESIGNRUPTC="'+UsDateTime(idate1900)+'" WHERE PCI_DATESIGNRUPTC IS NULL');
    end;
    if VersionBaseDest < '998.ZE' then
    begin
			ExecuteSQL('UPDATE CONTRATTRAVAIL SET PCI_TYPECDD="" WHERE PCI_TYPECDD IS NULL');
    end;
    if VersionBaseDest < '998.ZR' then
    begin
    	ExecuteSQL ('UPDATE CONTRATTRAVAIL SET PCI_SALAIREREF=0, PCI_ETABLIEUTRAV=PCI_ETABLISSEMENT, '+
      						'PCI_TRANSACTION="NON", PCI_PORTABPREV="NON", PCI_MOTIFEXCLUDSN = "", PCI_MOTIFSUSPAIE  = "", '+
                  'PCI_QUOTITE=0 WHERE PCI_SALAIREREF IS NULL');
      ExecuteSQL ('UPDATE CONTRATTRAVAIL SET PCI_REGIMELOCALDSN = "01" WHERE PCI_ETABLISSEMENT IN (SELECT ETB_ETABLISSEMENT From ETABCOMPL WHERE  ETB_REGIMEALSACE  = "X")');
      ExecuteSQL ('UPDATE CONTRATTRAVAIL SET PCI_REGIMELOCALDSN = "99" WHERE PCI_ETABLISSEMENT IN (SELECT ETB_ETABLISSEMENT From ETABCOMPL  WHERE  ETB_REGIMEALSACE  = "-")');
    end;
  end else if NomTable = 'DECLARATIONAED' then
  begin
    if VersionBaseDest < '998.Q' then
    begin
      ExecuteSQL('UPDATE DECLARATIONAED SET PDN_NBANNULEREMP=0 WHERE PDN_TYPEDECL="51" AND PDN_NBANNULEREMP IS NULL');
      ExecuteSQL('UPDATE DECLARATIONAED SET PDN_NBANNULEREMP=1 WHERE PDN_TYPEDECL="59" AND PDN_NBANNULEREMP IS NULL');
    end;
  end else if NomTable = 'PARPIECE' then
  begin
    if VersionBaseDest < '998.W' then
    begin
    	ExecuteSQL('UPDATE PARPIECE SET GPP_NUMMOISPIECE="-",GPP_NBCARAN=2,GPP_SEPNUMMENSUEL="-",GPP_CARSEPNUM="" WHERE GPP_NUMMOISPIECE IS NULL');
    end;
    if VersionBaseDest < '998.ZZ2' then
    begin
      ExecuteSQL('UPDATE parpiece SET GPP_STOCKSSDETAIL= "-" where GPP_STOCKSSDETAIL IS NULL');
    end;
    if VersionBaseDest < '998.ZZ2' then
    begin
      ExecuteSQL('UPDATE parpiece SET GPP_STOCKSSDETAIL= "-" where GPP_STOCKSSDETAIL IS NULL');
    end;
    if VersionBaseDest < '998.ZZM' then
    begin
      ExecuteSQL('UPDATE parpiece SET GPP_NUMAUTOLIG= "-" where GPP_NUMAUTOLIG IS NULL');
    end;
    if VersionBaseDest < '998.ZZX' then
    begin
      ExecuteSQL('UPDATE parpiece SET GPP_ECHGMESSOUT="",GPP_ECHGMESSIN="" where GPP_ECHGMESSOUT IS NULL');
    end;
    if VersionBaseDest < '998.ZZZD' then
    begin
      ExecuteSQL('UPDATE parpiece SET GPP_TAILLAGEAUTO="X" where GPP_TAILLAGEAUTO IS NULL');
    end;
  end else if NomTable = 'CODECPTA' then
  begin
    if VersionBaseDest < '998.W' then
    begin
    	ExecuteSQL('UPDATE CODECPTA SET GCP_VENTEACHAT="" WHERE GCP_VENTEACHAT IS NULL');
    end;
  end else if NomTable = 'GENERAUX' then
  begin
    if VersionBaseDest < '998.W' then
    begin
      ExecuteSQL('UPDATE GENERAUX SET G_TYPECPTTVA="TTC" WHERE G_NATUREGENE IN ("COC","COF","COS","COD","TIC","TID")');
      ExecuteSQL('UPDATE GENERAUX SET G_TYPECPTTVA="DIV" WHERE G_NATUREGENE IN ("BQE","CAI","DIV","EXT")');
      ExecuteSQL('UPDATE GENERAUX SET G_TYPECPTTVA="HT" WHERE G_NATUREGENE IN ("IMO","CHA","PRO")');
      ExecuteSQL('UPDATE GENERAUX SET G_TYPECPTTVA="TVA" WHERE G_GENERAL LIKE "445%"');
    end;
  end else if NomTable = 'PARFOU' then
  begin
    if VersionBaseDest < '998.W' then
    begin
  		ExecuteSql ('UPDATE PARFOU SET GRF_MULTIFOU="-" WHERE GRF_MULTIFOU IS NULL');
    end;
  end else if NomTable = 'DEPORTSAL' then
  begin
    if VersionBaseDest < '998.W' then
    begin
  		ExecuteSql ('UPDATE DEPORTSAL SET PSE_PRECOMPTEIJSS="-" , PSE_TYPMETIJSSMAL="DOS" , PSE_METHODIJSSMAL="" , PSE_TYPMETIJSSAT = "DOS", PSE_METHODIJSSAT="" WHERE PSE_PRECOMPTEIJSS IS NULL');
    end;
  end else if NomTable = 'ABSENCESALARIE' then
  begin
    if VersionBaseDest < '998.W' then
    begin
	  	ExecuteSql ('UPDATE ABSENCESALARIE SET PCN_MTIJSSBRUTE=0, PCN_MTIJSSBRUTE2=0, PCN_MTIJSSNETTE=0, PCN_MTIJSSNETTE2=0 , PCN_NBABSPRECOMPT=0, PCN_NBABSPRECOMPT2=0 WHERE PCN_MTIJSSBRUTE IS NULL');
  	  ExecuteSql ('UPDATE ABSENCESALARIE SET PCN_ASPROLONGATION = "-" WHERE PCN_ASPROLONGATION IS NULL');
    	ExecuteSql ('UPDATE ABSENCESALARIE SET PCN_DARRETINITIAL = "' + UsDateTime(Idate1900) + '" WHERE PCN_DARRETINITIAL IS NULL');
	  	ExecuteSql ('UPDATE ABSENCESALARIE SET PCN_TYPEIMPUTE="REL" WHERE PCN_TYPEMVT="CPA" AND PCN_TYPECONGE="PRI" AND PCN_PERIODEPAIE LIKE "Reliquat%" AND PCN_TYPEIMPUTE<>"REL" AND (PCN_TYPEIMPUTE IS NULL OR PCN_TYPEIMPUTE="")');
    end;
    if VersionBaseDest < '998.ZZ6' then
    begin
     ExecuteSQL ('update absencesalarie SET PCN_DSN="-"');
    end;
  // ------
  end else if NomTable = 'SALARIECOMPL' then
  begin
    if VersionBaseDest < '998.ZR' then
    begin
			ExecuteSQL('UPDATE SALARIESCOMPL SET PSZ_CTRLCIVILNIR = "-" WHERE PSZ_CTRLCIVILNIR IS NULL');
    end;
  // ------
  end else if NomTable = 'REGLTIJSS' then
  begin
    if VersionBaseDest < '998.W' then
    begin
  		ExecuteSql ('UPDATE REGLTIJSS SET PRI_PRECOMPTE="-", PRI_REGULARISATION ="-" WHERE PRI_PRECOMPTE IS NULL');
    end;
  end else if NomTable = 'DECLARANTATTEST' then
  begin
    if VersionBaseDest < '998.W' then
    begin
	  	ExecuteSql ('UPDATE DECLARANTATTEST SET PDA_EMAIL="" WHERE PDA_EMAIL IS NULL');
    end;
    if VersionBaseDest < '998.ZE' then
    begin
	    ExecuteSQL ('UPDATE DECLARANTATTEST SET PDA_NOMPERSONNE="" WHERE PDA_NOMPERSONNE IS NULL');
    end;
  end else if NomTable = 'ATTESTATIONS' then
  begin
    if VersionBaseDest < '998.W' then
    begin
  		ExecuteSql ('UPDATE ATTESTATIONS SET PAS_ATTESTEDI="-" , PAS_ATTESTRECTIFIC="-" , PAS_TYPEREPRISE="" , PAS_NUMDECLARANT = "" WHERE PAS_ATTESTEDI IS NULL');
    end;
  end else if NomTable = 'EMETTEURSOCIAL' then
  begin
    if VersionBaseDest < '998.W' then
    begin
	  	ExecuteSql ('UPDATE EMETTEURSOCIAL SET PET_NOMATTESTEDI="" , PET_PRENATTESTEDI="" , PET_EMAILATTESTEDI="" WHERE PET_NOMATTESTEDI IS NULL');
    end;
    if VersionBaseDest < '998.ZR' then
    begin
 			ExecuteSQL ('UPDATE EMETTEURSOCIAL SET PET_CONTACTDSN=PET_CONT1DUDS, PET_DOMAINEDSN=PET_DOMAINEDUDS1, '+
      						'PET_TELDSN=PET_TEL1DADSU, PET_FAXDSN=PET_FAX1DADSU, PET_MAILDSN=PET_APPEL1DUDS '+
                  'WHERE PET_CONTACTDSN IS NULL');
      ExecuteSQL ('UPDATE EMETTEURSOCIAL SET PET_CIVILDSN="01" WHERE PET_CIVIL1DADSU="MR"');
      ExecuteSQL ('UPDATE EMETTEURSOCIAL SET PET_CIVILDSN="02" WHERE PET_CIVIL1DADSU="MLE" OR PET_CIVIL1DADSU="MME"');
      ExecuteSQL ('UPDATE EMETTEURSOCIAL SET PET_CIVILDSN="" WHERE PET_CIVILDSN IS NULL');
      ExecuteSQL ('UPDATE EMETTEURSOCIAL SET PET_CONTACTDSN2=PET_CONT2DUDS, PET_DOMAINEDSN2=PET_DOMAINEDUDS2, PET_TELDSN2=PET_TEL2DADSU, PET_FAXDSN2=PET_FAX2DADSU, PET_MAILDSN2=PET_APPEL2DUDS WHERE PET_CONTACTDSN2 IS NULL');
      ExecuteSQL ('UPDATE EMETTEURSOCIAL SET PET_CIVILDSN2="01" WHERE PET_CIVIL2DADSU="MR"');
      ExecuteSQL ('UPDATE EMETTEURSOCIAL SET PET_CIVILDSN2="02" WHERE PET_CIVIL2DADSU="MLE" OR PET_CIVIL2DADSU="MME"');
      ExecuteSQL ('UPDATE EMETTEURSOCIAL SET PET_CIVILDSN2="" WHERE PET_CIVILDSN2 IS NULL');
      ExecuteSQL ('UPDATE EMETTEURSOCIAL SET PET_CONTACTDSN3=PET_CONT3DUDS, PET_DOMAINEDSN3=PET_DOMAINEDUDS3, PET_TELDSN3=PET_TEL3DADSU, PET_FAXDSN3=PET_FAX3DADSU, PET_MAILDSN3=PET_APPEL3DUDS WHERE PET_CONTACTDSN3 IS NULL');
      ExecuteSQL ('UPDATE EMETTEURSOCIAL SET PET_CIVILDSN3="01" WHERE PET_CIVIL3DADSU="MR"');
      ExecuteSQL ('UPDATE EMETTEURSOCIAL SET PET_CIVILDSN3="02" WHERE PET_CIVIL3DADSU="MLE" OR PET_CIVIL3DADSU="MME"');
      ExecuteSQL ('UPDATE EMETTEURSOCIAL SET PET_CIVILDSN3="" WHERE PET_CIVILDSN3 IS NULL');
    end;
  end else if NomTable = 'METHCALCULSALMOY' then
  begin
    if VersionBaseDest < '998.W' then
    begin
  		ExecuteSql ('UPDATE METHCALCULSALMOY SET PSM_TYPECALC="IDR" WHERE PSM_TYPECALC IS NULL');
    end;
  end else if NomTable = 'BTTYPEAFFAIRE' then
  begin
    if VersionBaseDest < '998.W' then
    begin
    	ExecuteSql ('UPDATE BTTYPEAFFAIRE SET BTY_PRIOCONTRAT="" WHERE BTY_PRIOCONTRAT IS NULL');
    end;
  end else
  if NomTable = 'ARTICLE' then
  begin
    if VersionBaseDest < '998.ZZS' then
    begin
      ExecuteSql ('UPDATE ARTICLE SET GA_RELIQUATMT="-" WHERE GA_RELIQUATMT IS NULL');
    end;
    if VersionBaseDest < '998.W' then
    begin
    	ExecuteSql ('UPDATE ARTICLE SET GA_GEREDEMPRIX="X" WHERE GA_GEREDEMPRIX IS NULL');
    end;
    if VersionBaseDest < '998.Z7' then
    begin
    	ExecuteSql ('UPDATE ARTICLE SET GA_SECTION="",GA_GEREANAL="X" WHERE GA_SECTION IS NULL');
    end;
    if VersionBaseDest < '998.ZZ9' then
    begin
    	ExecuteSql ('UPDATE ARTICLE SET GA_QUALIFUNITEACH="",GA_PAUA=0, GA_COEFCONVQTEACH=0 WHERE GA_QUALIFUNITEACH IS NULL');
    end;
  end else
  if NomTable = 'BDETETUDE' then
  begin
    if VersionBaseDest < '998.W' then
    begin
    	ExecuteSql ('UPDATE BDETETUDE SET BDE_FOURNISSEUR="" WHERE BDE_FOURNISSEUR IS NULL');
    end;
  end else
  if NomTable = 'CODECPTA' then
  begin
    if VersionBaseDest < '998.W' then
    begin
    	ExecuteSql ('UPDATE CODECPTA SET GCP_FAMILLETAXE="" WHERE GCP_FAMILLETAXE IS NULL');
    end;
  end else
  if NomTable = 'DECISIONACHLIG' then
  begin
    if VersionBaseDest < '998.W' then
    begin
      ExecuteSql ('UPDATE DECISIONACHLIG SET BAD_PABASE=0 WHERE BAD_PABASE IS NULL');
      ExecuteSql ('UPDATE DECISIONACHLIG SET BAD_COEFUAUS=0 WHERE BAD_COEFUAUS IS NULL');
      ExecuteSql ('UPDATE DECISIONACHLIG SET BAD_COEFUSUV=0 WHERE BAD_COEFUSUV IS NULL');
    end;
  end else
  if NomTable = 'DECISIONACHLFOU' then
  begin
    if VersionBaseDest < '998.W' then
    begin
    	ExecuteSql ('UPDATE DECISIONACHLFOU SET BDF_COEFUAUS=0 WHERE BDF_COEFUAUS IS NULL');
    end;
  end else if NomTable = 'LIGNEOUVPLAT' then
  begin
    if VersionBaseDest < '998.W' then
    begin
      UpdateDecoupeLigneOuvPlat('BOP_MONTANTPR=0, BOP_COEFFG=0, BOP_COEFFR=0, BOP_COEFFC=0',' AND (BOP_COEFFC IS NULL)');
      UpdateDecoupeLigneOuvPlat('BOP_NATURETRAVAIL="", BOP_FOURNISSEUR=""',' AND (BOP_NATURETRAVAIL IS NULL)');
    end;
    if VersionBaseDest < '998.ZZ2' then
    begin
      UpdateDecoupeLigneOuvPlat('BOP_TENUESTOCK=(select ga_tenuestock from article where ga_article=bop_article),BOP_COEFCONVQTE=0,BOP_COEFCONVQTEVTE=0',' AND (bop_tenuestock IS NULL)');
    end;
  end else if NomTable = 'AFFAIRE' then
  begin
    if VersionBaseDest < '998.W' then
    begin
      ExecuteSql ('UPDATE AFFAIRE SET AFF_FACTURE=(SELECT ##TOP 1## T_FACTURE FROM TIERS WHERE T_TIERS=AFF_TIERS AND T_NATUREAUXI IN ("CLI","PRO")) WHERE AFF_FACTURE IS NULL');
      ExecuteSql ('UPDATE AFFAIRE SET AFF_MANDATAIRE="",AFF_CODEBQ="",AFF_TYPEPAIE="",AFF_BQMANDATAIRE="",AFF_IDSPIGAO=0 WHERE AFF_MANDATAIRE IS NULL');
    end;
    if VersionBaseDest < '998.Z3' then
    begin
      ExecuteSql ('UPDATE AFFAIRE SET AFF_MODELEWORD="" WHERE AFF_MODELEWORD IS NULL');
    end;
    if VersionBaseDest < '998.ZE' then
    begin
      ExecuteSql ('UPDATE AFFAIRE SET AFF_SSTRAITANCE="-", AFF_DATESSTRAIT="'+USDATETIME(IDate1900)+'" where aff_sstraitance is null');
    end;
    if VersionBaseDest < '998.ZZ7' then
    begin
      ExecuteSql ('UPDATE AFFAIRE SET AFF_DEMANDEREGLE="-",AFF_TVAREDUITE="-" WHERE AFF_DEMANDEREGLE IS NULL');
    end;
    if VersionBaseDest < '998.ZZZL' then
    begin
      ExecuteSql ('UPDATE AFFAIRE SET AFF_COEFFG=0,AFF_COEFFS=0,AFF_COEFSAV=0,AFF_COEFFD=0,AFF_TAUXRG=0,AFF_DATERG="'+USDATETIME(Idate1900)+'",AFF_MTAVANCE=0,AFF_DEBRESTAVANCE=0,AFF_FINRESTAVANCE=0,AFF_COEFMARG=0,AFF_SECTEUR=""  WHERE AFF_COEFFG IS NULL');
    end;
  end else if NomTable = 'FACTAFF' then
  begin
    if VersionBaseDest < '998.W' then
    begin
    	ExecuteSql ('UPDATE FACTAFF SET AFA_DATEDEBUTFAC="'+USDATETIME(IDate1900)+'",AFA_DATEFINFAC="'+USDATETIME(Idate1900)+'" WHERE AFA_DATEDEBUTFAC IS NULL');
    end;
  end else if NomTable = 'HRPARAMPLANNING' then
  begin
    if VersionBaseDest < '998.W' then
    begin
	    ExecuteSql ('UPDATE HRPARAMPLANNING SET HPP_REGCOL1="-", HPP_REGCOL2="-", HPP_REGCOL3="-" WHERE HPP_REGCOL1 IS NULL');
    end;
    if VersionBaseDest < '998.ZZF' then
    begin
	    ExecuteSql ('UPDATE HRPARAMPLANNING SET HPP_AFFEVTMAT="-" WHERE HPP_AFFEVTMAT IS NULL');
    end;
    if VersionBaseDest < '998.ZZQ' then
    begin
	    ExecuteSql ('UPDATE HRPARAMPLANNING SET HPP_AFFEVTCHA="-",HPP_AFFEVTINT="X",HPP_AFFABSSAL="X", HPP_AFFACTGRC="X" WHERE HPP_AFFEVTCHA IS NULL');
    end;
  end else if nomTable ='PARCTIERS' then
  begin
    if VersionBaseDest < '998.W' then
    begin
    	ExecuteSql ('UPDATE PARCTIERS SET BP1_DATEFINSERIA="20991231" WHERE BP1_DATEFINSERIA IS NULL');
    end;
  end else if nomTable ='PIEDECHE' then
  begin
    if VersionBaseDest < '998.W' then
    begin
    	ExecuteSql ('UPDATE PIEDECHE SET GPE_CODE="",GPE_TIMBRE=0 WHERE GPE_CODE IS NULL');
    end;
    if VersionBaseDest < '998.X' then
    begin
    	ExecuteSql ('UPDATE PIEDECHE SET GPE_FOURNISSEUR="" WHERE GPE_FOURNISSEUR IS NULL');
    end;
  end else if nomTable ='LIGNE' then
  begin
    if VersionBaseDest < '998.W' then
    begin
      UpdateDecoupeLigne('GL_IDSPIGAO=""',' AND GL_IDSPIGAO IS NULL');
      UpdateDecoupeLigne('GL_COEFCONVQTEVTE=0',' AND GL_COEFCONVQTEVTE IS NULL');
    end;
    if VersionBaseDest < '998.ZZS' then
    begin
      UpdateDecoupeLigne('GL_MTRESTE=0',   ' AND GL_MTRESTE IS NULL');
      UpdateDecoupeLigne('GL_MTRELIQUAT=0',' AND GL_MTRELIQUAT IS NULL');
    end;
    if VersionBaseDest < '998.ZZZ' then
    begin
      UpdateDecoupeLigne('GL_GESTIONPTC="-"',   ' AND GL_GESTIONPTC IS NULL');
    end;
    if VersionBaseDest < '998.ZZZD' then
    begin
      UpdateDecoupeLigne('GL_QTESAIS=GL_QTEFACT, GL_RENDEMENT=0, GL_PERTE=0',' AND GL_QTESAIS IS NULL');
    end;
    if VersionBaseDest < '998.ZZZM' then
    begin
      UpdateDecoupeLigne('GL_CODEMARCHE=""',' AND GL_CODEMARCHE IS NULL');
    end;
  end else if nomTable ='LIGNEBASE' then
  begin
    if VersionBaseDest < '998.W' then
    begin
			UpdateDecoupeLigneBase('BLB_FOURNISSEUR="",BLB_BASEACHAT=0,BLB_VALEURACHAT=0,BLB_TYPEINTERV=""',' AND BLB_FOURNISSEUR IS NULL');
    end;
  end else if nomTable ='LIGNECOMPL' then
  begin
    if VersionBaseDest < '998.W' then
    begin
			UpDateDecoupeLigneCompl('GLC_VOIRDETAIL="-"',' AND GLC_VOIRDETAIL IS NULL');
    end;
    if VersionBaseDest < '998.ZZF' then
    begin
			UpDateDecoupeLigneCompl('GLC_CODEMATERIEL="",GLC_BTETAT="",GLC_IDEVENTMAT=0 ',' AND GLC_CODEMATERIEL IS NULL');
    end;
    if VersionBaseDest < '998.ZZM' then
    begin
			UpDateDecoupeLigneCompl('GLC_NUMEROTATION="",GLC_NUMFORCED="-" ',' AND GLC_NUMEROTATION IS NULL');
    end;
    if VersionBaseDest < '998.ZZZD' then
    begin
			UpDateDecoupeLigneCompl('GLC_MTMOPA=0,GLC_MTFOUPA=0,GLC_MTINTPA=0,GLC_MTLOCPA=0,GLC_MTMATPA=0,GLC_MTOUTPA=0,GLC_MTSTPA=0,GLC_MTAUTPA=0, '+
                              'GLC_MTMOPR=0,GLC_MTFOUPR=0,GLC_MTINTPR=0,GLC_MTLOCPR=0,GLC_MTMATPR=0,GLC_MTOUTPR=0,GLC_MTSTPR=0,GLC_MTAUTPR=0, '+
                              'GLC_MTMOPV=0,GLC_MTFOUPV=0,GLC_MTINTPV=0,GLC_MTLOCPV=0,GLC_MTMATPV=0,GLC_MTOUTPV=0,GLC_MTSTPV=0,GLC_MTAUTPV=0',' AND GLC_MTMOPA IS NULL');
    end;
    if VersionBaseDest < '998.ZZZK' then
    begin
      UpDateDecoupeLigneCompl('GLC_COEFCOND=0 ',' AND GLC_COEFCOND IS NULL');
    end;

    if VersionBaseDest < '998.ZZZY' then
    begin
      UpDateDecoupeLigneCompl('GLC_VISIBLE="X",GLC_SAISISSABLE="X" ',' AND GLC_VISIBLE IS NULL');
    end;

  end else if nomTable ='LIGNEOUV' then
  begin
    if VersionBaseDest < '998.W' then
    begin
			UpDateDecoupeLigneOUV('BLO_UNIQUEBLO=0,BLO_NATURETRAVAIL=""',' AND BLO_UNIQUEBLO IS NULL');
    end;
    if VersionBaseDest < '998.ZZ2' then
    begin
			UpDateDecoupeLigneOUV('BLO_COEFCONVQTE=0,BLO_COEFCONVQTEVTE=0',' AND BLO_COEFCONVQTE IS NULL');
    end;
    if VersionBaseDest < '998.ZZZD' then
    begin
			UpDateDecoupeLigneOUV('BLO_MTMOPA=0,BLO_MTFOUPA=0,BLO_MTINTPA=0,BLO_MTLOCPA=0,BLO_MTMATPA=0,BLO_MTOUTPA=0,BLO_MTSTPA=0,BLO_MTAUTPA=0,'+
                            'BLO_MTMOPR=0,BLO_MTFOUPR=0,BLO_MTINTPR=0,BLO_MTLOCPR=0,BLO_MTMATPR=0,BLO_MTOUTPR=0,BLO_MTSTPR=0,BLO_MTAUTPR=0,'+
                            'BLO_MTMOPV=0,BLO_MTFOUPV=0,BLO_MTINTPV=0,BLO_MTLOCPV=0,BLO_MTMATPV=0,BLO_MTOUTPV=0,BLO_MTSTPV=0,BLO_MTAUTPV=0,'+
                            'BLO_QTESAIS=BLO_QTEFACT,BLO_RENDEMENT=0,BLO_PERTE=0',' AND BLO_MTMOPA IS NULL');
    end;
    if VersionBaseDest < '998.ZZZM' then
    begin
			UpDateDecoupeLigneOUV('BLO_CODEMARCHE=""',' AND BLO_CODEMARCHE IS NULL');
    end;
  end else if nomTable ='LIGNEFAC' then
  begin
    if VersionBaseDest < '998.W' then
    begin
    	UpDateDecoupeLignefac('BLF_NATURETRAVAIL="",BLF_FOURNISSEUR="",BLF_UNIQUEBLO=0,BLF_TOTALTTCDEV=0',' AND BLF_NATURETRAVAIL IS NULL');
    end;
    if VersionBaseDest < '998.ZZ7' then
    begin
    	UpDateDecoupeLignefac('BLF_MTPRODUCTION=0',' AND BLF_MTPRODUCTION IS NULL');
    end;
    if VersionBaseDest < '998.ZZZM' then
    begin
    	UpDateDecoupeLignefac('BLF_CODEMARCHE=""',' AND BLF_CODEMARCHE IS NULL');
    end;
  end else if nomTable ='PIECE' then
  begin
    if VersionBaseDest < '998.W' then
    begin
    	UpDateDecoupePiece('GP_UNIQUEBLO=0',' AND GP_UNIQUEBLO IS NULL');
    end;
    if VersionBaseDest < '998.ZZF' then
    begin
    	UpDateDecoupePiece('GP_CODEMATERIEL="",GP_BTETAT="" ',' AND GP_CODEMATERIEL IS NULL');
    end;
    if VersionBaseDest < '998.ZZS' then
    begin
    	UpDateDecoupePiece('GP_ATTACHEMENT=""',' AND GP_ATTACHEMENT IS NULL');
    end;
    if VersionBaseDest < '998.ZZV' then
    begin
    	UpDateDecoupePiece('GP_AUTOLIQUID="-"',' AND GP_AUTOLIQUID IS NULL');
    end;
    if VersionBaseDest < '998.ZZZD' then
    begin
    	UpDateDecoupePiece('GP_TOTALMOPA=0, GP_TOTALFOUPA=0, GP_TOTALINTPA=0,GP_TOTALLOCPA=0, GP_TOTALMATPA=0,GP_TOTALOUTPA=0, GP_TOTALSTPA=0, GP_TOTALAUTPA=0,'+
                         'GP_TOTALMOPR=0, GP_TOTALFOUPR=0, GP_TOTALINTPR=0,GP_TOTALLOCPR=0, GP_TOTALMATPR=0,GP_TOTALOUTPR=0, GP_TOTALSTPR=0, GP_TOTALAUTPR=0,'+
                         'GP_TOTALMOPV=0, GP_TOTALFOUPV=0, GP_TOTALINTPV=0,GP_TOTALLOCPV=0, GP_TOTALMATPV=0,GP_TOTALOUTPV=0, GP_TOTALSTPV=0, GP_TOTALAUTPV=0',' AND GP_TOTALMOPA IS NULL');
    end;
    if VersionBaseDest < '998.ZZZL' then
    begin
    	UpDateDecoupePiece('GP_BSVREF=""',' AND GP_BSVREF IS NULL');
    end;
    if VersionBaseDest < '998.ZZZZ1' then
    begin
    	UpDateDecoupePiece('GP_HORSCOMPTA="-"',' AND GP_HORSCOMPTA IS NULL');
    end;
  end else if nomTable ='PIECETRAIT' then
  begin
    if VersionBaseDest < '998.W' then
    begin
	    ExecuteSql ('UPDATE PIECETRAIT SET BPE_REGLSAISIE="-" WHERE BPE_REGLSAISIE IS NULL');
    end;
  end else if nomTable ='PIEDBASE' then
  begin
    if VersionBaseDest < '998.W' then
    begin
    	UpDateDecoupePiedbase('GPB_FOURN="",GPB_BASEACHAT=0,GPB_VALEURACHAT=0,GPB_DELTAACHAT=0,GPB_TYPEINTERV=""',' AND GPB_FOURN IS NULL');
    end;
  end else if nomTable ='PIECERG' then
  begin
    if VersionBaseDest < '998.W' then
    begin
    	ExecuteSQL('UPDATE PIECERG SET PRG_APPLICABLE="X",PRG_FOURN="",PRG_NUMERORIB=0 WHERE PRG_APPLICABLE IS NULL');
    end;
  end else if NomTable = 'REMUNERATION' then
  begin
    if VersionBaseDest < '998.W' then
    begin
      ExecuteSql ('UPDATE REMUNERATION SET PRM_CONDIUTILIS1="" WHERE PRM_CONDIUTILIS1 IS NULL');
      ExecuteSql ('UPDATE REMUNERATION SET PRM_CONDIUTIL2="" WHERE PRM_CONDIUTIL2 IS NULL');
      ExecuteSql ('UPDATE REMUNERATION SET PRM_VALELTDYNTAB="" WHERE PRM_VALELTDYNTAB IS NULL');
    end;
    if VersionBaseDest < '998.ZR' then
    begin
    	ExecuteSql ('UPDATE REMUNERATION SET PRM_PERIODERUB="100", PRM_ISRUBPERIOD="-", PRM_NATUREPRIME="", '+
      						'PRM_APPLICNATURE="'+UsDateTime (Idate1900)+'", '+
                  'PRM_MODIFNATURE="'+UsDateTime (Idate1900)+'", '+
                  'PRM_AUTREELEMENT = "", PRM_APPLICAUTREELE = "'+UsDateTime (Idate1900)+'", '+
                  'PRM_MODIFAUTREELEM = "'+UsDateTime (Idate1900)+'" WHERE PRM_PERIODERUB IS NULL');
    end;
  end else if NomTable = 'DECLARATIONS' then
  begin
    if VersionBaseDest < '998.W' then
    begin
      ExecuteSql ('UPDATE DECLARATIONS SET PDT_CIRCACC7BIS="",'+
                  ' PDT_ACCIDENTTRAV="-", PDT_ACCIDENTTRAJET="-",'+
                  ' PDT_CONTRATCDI="-", PDT_CONTRATCDD="-",'+
                  ' PDT_CONTRATAPP="-", PDT_CONTRATINT="-",'+
                  ' PDT_CONTRATAUTRE="-", PDT_BOOLTEMOIN="-",'+
                  ' PDT_PERSONNEAVISEE="-" WHERE PDT_CIRCACC7BIS IS NULL');
    end;
  end else if NomTable = 'SALARIESCOMPL' then
  begin
    if VersionBaseDest < '998.W' then
    begin
      ExecuteSql ('INSERT INTO SALARIESCOMPL (PSZ_SALARIE) SELECT PSA_SALARIE '+
                  'FROM SALARIES WHERE PSA_SALARIE NOT IN (SELECT PSZ_SALARIE FROM SALARIESCOMPL)');
      ExecuteSql ('UPDATE SALARIESCOMPL SET PSZ_TYPMODEARRONDI = "ETB" '+
                  'WHERE (PSZ_TYPMODEARRONDI IS NULL OR PSZ_TYPMODEARRONDI="") AND PSZ_SALARIE<>""');
      ExecuteSql ('UPDATE SALARIESCOMPL SET PSZ_TYPMODEARRSLDC = "ETB" '+
                  'WHERE (PSZ_TYPMODEARRSLDC IS NULL OR PSZ_TYPMODEARRSLDC="") AND PSZ_SALARIE<>""');
      ExecuteSql ('UPDATE SALARIESCOMPL SET PSZ_CPMODEARRSLDC = (SELECT ETB_CPMODEARRSLDC FROM ETABCOMPL '+
                  'WHERE ETB_ETABLISSEMENT =(SELECT PSA_ETABLISSEMENT FROM SALARIES WHERE PSA_SALARIE = PSZ_SALARIE)) '+
                  'WHERE (PSZ_CPMODEARRSLDC IS NULL OR PSZ_CPMODEARRSLDC="") AND PSZ_SALARIE<>""');
      ExecuteSql ('UPDATE SALARIESCOMPL SET PSZ_CPMODEARRONDI = (SELECT ETB_CPMODEARRONDI FROM ETABCOMPL '+
                  'WHERE ETB_ETABLISSEMENT =(SELECT PSA_ETABLISSEMENT FROM SALARIES WHERE PSA_SALARIE = PSZ_SALARIE)) '+
                  'WHERE (PSZ_CPMODEARRONDI IS NULL OR PSZ_CPMODEARRONDI="") AND PSZ_SALARIE<>""');
    end;
    if VersionBaseDest < '998.ZE' then
    begin
      ExecuteSQL('UPDATE SALARIESCOMPL SET PSZ_DIGITALELECTRO="" WHERE PSZ_DIGITALELECTRO IS NULL');
      ExecuteSQL('UPDATE SALARIESCOMPL SET PSZ_DIGITALPAPER="" WHERE PSZ_DIGITALPAPER IS NULL');
      ExecuteSql('UPDATE SALARIESCOMPL SET PSZ_DIGITALEXCLU="" WHERE PSZ_DIGITALEXCLU IS NULL');
    end;
  end else if nomTable ='CONSOMMATIONS' then
  begin
    if VersionBaseDest < '998.ZT' then
    begin
      ExecuteSql ('UPDATE CONSOMMATIONS SET BCO_LINKEQUIPE="" WHERE BCO_LINKEQUIPE IS NULL');
    end;
    if VersionBaseDest < '998.ZZZL' then
    begin
      ExecuteSql('update consommations set '+
                 'bco_familleniv1=(select ga_familleniv1 from article where ga_article=bco_article), '+
                 'bco_familleniv2=(select ga_familleniv2 from article where ga_article=bco_article), '+
                 'bco_familleniv3=(select ga_familleniv3 from article where ga_article=bco_article) '+
                 'where bco_familleniv1 is null');
    end;
    if VersionBaseDest < '998.ZZ4' then
    begin
      ExecuteSql ('UPDATE CONSOMMATIONS SET BCO_FAMILLETAXE1="" WHERE BCO_FAMILLETAXE1 IS NULL');
      ExecuteSql ('UPDATE CONSOMMATIONS SET BCO_DATETRAVAUX=BCO_DATEMOUV+" 08:00:00" WHERE BCO_DATETRAVAUX IS NULL');
    end;
    if VersionBaseDest < '998.ZZZ' then
    begin
      ExecuteSql ('UPDATE CONSOMMATIONS SET BCO_FROMDEVIS="-" WHERE BCO_FROMDEVIS IS NULL');
    end;
    if VersionBaseDest < '998.ZZZP' then
    begin
      SQL := 'UPDATE CONSOMMATIONS SET '+
             'BCO_FAMILLENIV1=(SELECT GA_FAMILLENIV1 FROM ARTICLE WHERE GA_ARTICLE=BCO_ARTICLE), '+
             'BCO_FAMILLENIV2=(SELECT GA_FAMILLENIV2 FROM ARTICLE WHERE GA_ARTICLE=BCO_ARTICLE), '+
             'BCO_FAMILLENIV3=(SELECT GA_FAMILLENIV3 FROM ARTICLE WHERE GA_ARTICLE=BCO_ARTICLE) '+
             'WHERE BCO_FAMILLENIV1 IS NULL';
      ExecuteSql (SQL);
    end;
  end
  else if nomTable ='TIERS' then
  begin
    if VersionBaseDest < '998.ZZ8' then
    begin
      if ISColumnExists ('T_CODEBARRE','TIERS') and ISColumnExists ('T_QUALIFCODEBARRE','TIERS') then
      begin
      ExecuteSql ('UPDATE TIERS SET T_CODEBARRE="",T_QUALIFCODEBARRE="" WHERE T_CODEBARRE IS NULL');
    end;
    end;
  end
  else if NomTable = 'BTFAMILLEMAT' then
  begin
    if VersionBaseDest < '998.ZZI' then
    begin
      ExecuteSql ('UPDATE BTFAMILLEMAT SET BFM_NONGEREPLANNING="X" WHERE BFM_NONGEREPLANNING IS NULL');
    end
  end
  else if NomTable = 'HRPARAMPLANNING' then
  begin
    if VersionBaseDest < '998.ZZZ4' then
    begin
      ExecuteSql ('UPDATE HRPARAMPLANNING SET HPP_AFFEVTCHA="-", HPP_AFFEVTINT="-", HPP_AFFABSSAL="-", HPP_AFFACTGRC="-"');
    end;
  end
  else if NomTable = 'SOUCHE' then
  begin
    if VersionBaseDest < '998.ZZL' then
    begin
      if ISColumnExists ('SH_NUMMOISPIECE','SOUCHE') then
      begin
      ExecuteSql ('update SOUCHE SET SH_NUMMOISPIECE="-" WHERE SH_NUMMOISPIECE IS NULL');
      ExecuteSql ('update SOUCHE SET SH_NUMMOISPIECE=(select ##TOP 1## GPP_NUMMOISPIECE FROM PARPIECE where GPP_SOUCHE=SH_SOUCHE) WHERE SH_TYPE="GES"');
    end;
    end;
  end
  else if NomTable = 'BTPARDOC' then
  begin
    if VersionBaseDest < '998.ZZS' then
    begin
      ExecuteSql ('update BTPARDOC SET BPD_IMPDETAVE="-", BPD_IMPCUMULAVE="-"  WHERE BPD_IMPDETAVE IS NULL');
    end;
  end else if NomTable ='NOMENLIG' then
  begin
    if VersionBaseDest < '998.ZZZD' then
    begin
      ExecuteSql ('UPDATE NOMENLIG SET GNL_PERTE=0,GNL_RENDEMENT=0,GNL_QTESAIS=GNL_QTE,GNL_QUALIFUNITEACH="" WHERE GNL_PERTE IS NULL');
    end;
  end
  else if NomTable = 'COMMERCIAL' then
  begin
    if VersionBaseDest < '998.ZZZX' then
    begin
      ExecuteSql ('UPDATE COMMERCIAL SET GCL_FERME="-" WHERE GCL_FERME IS NULL');
    end;
  end;


end;

function TMajStructBTP.IsATraiter(NomTable: string;Force : boolean): boolean;
var QQ : TQUery;
    Versionref,Versiondest : integer;
begin
  if (Force) or (IsModeMajHalley) then
  begin
    Result := true;
    exit;
  end;
  QQ := OpenSQLDb (DBREF,'SELECT DT_NUMVERSION FROM DETABLES WHERE DT_NOMTABLE="'+NomTable+'"',true);
  VersionRef := QQ.findfield('DT_NUMVERSION').AsInteger;
  ferme(QQ);
  QQ := OpenSQLDb (DBSOC,'SELECT DT_NUMVERSION FROM DETABLES WHERE DT_NOMTABLE="'+NomTable+'"',true);
  VersionDest := QQ.findfield('DT_NUMVERSION').AsInteger;
  ferme(QQ);
  result:= (VersionRef > Versiondest);
end;

procedure TMajStructBTP.SetCodeStatus(const Value: integer);
begin
  fCodeStatus := Value;
end;

function TMajStructBTP.GetRequete(TheListRef : TOB;DBREF: TDatabase) : string;
begin
  if (Pos (';',TheListRef.GetValue('BTV_NOMELT'))> 0) or IsNumeric (TheListRef.GetValue('BTV_NOMELT')) then
  begin
    // On a affaire ici a une branche
    result := 'SELECT * FROM PARAMSOC WHERE SOC_TREE LIKE "'+TheListRef.GetValue('BTV_NOMELT')+'%"'
  end else
  begin
    // On a affaire ici a un parametre nommé
    result := 'SELECT * FROM PARAMSOC WHERE SOC_NOM = "'+TheListRef.GetValue('BTV_NOMELT')+'"';
  end;
end;

procedure TMajStructBTP.GetParamSocs(TheListRef,TheParamSocsRef,TheParamSocsDest: TOB);
var Sql : string;
    QQ : TQUery;
begin
  Sql := GetRequete(TheListRef,DBREF);
  QQ := OpenSQLDb(DBREF,Sql,True);
  TheParamSocsRef.LoadDetailDB ('PARAMSOC','','',QQ,false);
  ferme(QQ);
  Sql := GetRequete(TheListRef,DBSOC);
  QQ := OpenSql (Sql,True);
  TheParamSocsDest.LoadDetailDB ('PARAMSOC','','',QQ,false);
  ferme(QQ);
  //
end;

procedure TMajStructBTP.GetAllParamSocs(TheParamSocsRef,TheParamSocsDest: TOB);
var Sql : string;
    QQ : TQUery;
begin
  Sql := 'SELECT * FROM PARAMSOC';
  QQ := OpenSQLDb(DBREF,Sql,True);
  TheParamSocsRef.LoadDetailDB ('PARAMSOC','','',QQ,false);
  ferme(QQ);
  //
  QQ := OpenSql (Sql,True);
  TheParamSocsDest.LoadDetailDB ('PARAMSOC','','',QQ,false);
  ferme(QQ);
  //
end;

procedure TMajStructBTP.DropMenuDest(NumeroMenu : string);
var Sql,Menus : string;
    Mn1,Mn2,Mn3,Mn4 : string;
begin
	menus := NumeroMenu;
  Mn1 := '';
  Mn2 := '';
  Mn3 := '';
  Mn4 := '';
  // ----------------
  Mn1 := READTOKENST(Menus);
  if Menus <> '' then
  begin
  	Mn2 := READTOKENST(Menus);
  end;
  if Menus <> '' then
  begin
  	Mn3 := READTOKENST(Menus);
  end;
  if Menus <> '' then
  begin
  	Mn4 := READTOKENST(Menus);
  end;
  Sql := 'DELETE FROM MENU WHERE MN_1='+Mn1;
  if Mn2 <> '' then
  begin
  	Sql := Sql + ' AND MN_2='+Mn2;
  end;
  if Mn3 <> '' then
  begin
  	Sql := Sql + ' AND MN_3='+Mn3;
  end;
  if Mn4 <> '' then
  begin
  	Sql := Sql + ' AND MN_4='+Mn4;
  end;
//  Sql := SQl + ' AND MN_TAG < 900000';
  TRY
    ExecuteSql (Sql);
  except
  ENd;
  //
  TRY
    ExecuteSql ('DELETE FROM MENU WHERE MN_1=0 AND MN_2='+Mn1); //+' AND MN_TAG < 900000');
  EXCEPT
  END;
end;

procedure TMajStructBTP.GetMenus(NumeroMenu : String;TheMenusRef,TheMenusDest: TOB);
var Sql,Menus : string;
    QQ : TQUery;
    Mn1,Mn2,Mn3,Mn4 : string;
begin
	menus := NumeroMenu;
  Mn1 := '';
  Mn2 := '';
  Mn3 := '';
  Mn4 := '';
  //
  Mn1 := READTOKENST(Menus);
  if Menus <> '' then
  begin
  	Mn2 := READTOKENST(Menus);
  end;
  if Menus <> '' then
  begin
  	Mn3 := READTOKENST(Menus);
  end;
  if Menus <> '' then
  begin
  	Mn4 := READTOKENST(Menus);
  end;
  // --- Pour recuperer le niveau 0 du menu
  Sql := 'SELECT * FROM MENU WHERE MN_1=0 AND MN_2='+Mn1+' AND MN_TAG < 900000';
  QQ := OpenSQLDb(DBREF,Sql,True);
  TheMenusRef.LoadDetailDB ('MENU','','',QQ,false);
  ferme(QQ);
  QQ := OpenSql (Sql,True);
  TheMenusDest.LoadDetailDB ('MENU','','',QQ,false);
  ferme(QQ);
  // ---
  Sql := 'SELECT * FROM MENU WHERE MN_1='+Mn1;
  if Mn2 <> '' then
  begin
  	Sql := Sql + ' AND MN_2='+Mn2;
  end;
  if Mn3 <> '' then
  begin
  	Sql := Sql + ' AND MN_3='+Mn3;
  end;
  if Mn4 <> '' then
  begin
  	Sql := Sql + ' AND MN_4='+Mn4;
  end;
  QQ := OpenSQLDb(DBREF,Sql+' ORDER BY MN_1,MN_TAG',True);
  TheMenusRef.LoadDetailDB ('MENU','','',QQ,true);
  ferme(QQ);
  //
  Sql := SQl + ' AND MN_TAG < 900000 ORDER BY MN_1,MN_TAG';
  QQ := OpenSql (Sql,True);
  TheMenusDest.LoadDetailDB ('MENU','','',QQ,true);
  ferme(QQ);
  //
end;

procedure TMajStructBTP.UpDateDecoupeLigneOuvPlat(SetSQL: string; SetCondition: string='');
var NMin, NMax, NCount, NN1, NN2, NbStep: integer;
  ListeNat, ListeSouche: TStrings;
  Q: TQuery;
  QMax: TQuery;
  i, k: integer;
  NbDebutTranche: integer;
  NbFinTranche: integer;
  //PremierNumero: integer;
  DernierNumero: integer;
  //SQL : string ;
begin
  // Lecture des natures
  ListeNat := TStringList.Create;
  ListeSouche := TStringList.Create;
  Q := OpenSQL('Select DISTINCT BOP_NATUREPIECEG, BOP_SOUCHE FROM LIGNEOUVPLAT', True);
  while not Q.EOF do
  begin
    ListeNat.Add(Q.Fields[0].AsString);
    ListeSouche.Add(Q.Fields[1].AsString);
    Q.Next;
  end;
  Ferme(Q);
  // Balayage des pièces avec découpe
  for i := 0 to ListeNat.Count - 1 do
  begin
    // Récupération du plus grand numéros possible
    QMax := OpenSQL('SELECT MIN(BOP_NUMERO), MAX(BOP_NUMERO),COUNT(*) FROM LIGNEOUVPLAT WHERE BOP_NATUREPIECEG="' + ListeNat[i] + '" AND BOP_SOUCHE="' + ListeSouche[i] +
                    '"', True);
    //PremierNumero := QMax.Fields[0].AsInteger;
    DernierNumero := QMax.Fields[1].AsInteger;
    NCount := QMax.Fields[2].AsInteger;
    Ferme(QMax);

    if NCount > 50000 then
    begin
      // Première tranche traitée : 0 à 1000000
      NbDebutTranche := 0;
      NbFinTranche := 1000000;

      while (NbDebutTranche < DernierNumero) do
      begin
        // Récupère le premier et le dernier numéro à traiter dans la tranche
        QMax := OpenSQL('SELECT MIN(BOP_NUMERO), MAX(BOP_NUMERO),COUNT(*) FROM LIGNEOUVPLAT WHERE BOP_NATUREPIECEG="' + ListeNat[i] + '" AND BOP_SOUCHE="' +
          ListeSouche[i] + '" AND BOP_NUMERO>=' + IntToStr(NbDebutTranche) + ' AND BOP_NUMERO<' + IntToStr(NbFinTranche) + '', True);
        NMin := QMax.Fields[0].AsInteger;
        NMax := QMax.Fields[1].AsInteger;
        NCount := QMax.Fields[2].AsInteger;
        Ferme(QMax);

        if NCount > 0 then
        begin
          NbStep := (NMax - NMin) div 25000;
          for k := 0 to NbStep do
          begin
            NN1 := NMin + k * 25000;
            if k < NbStep then NN2 := NN1 + 25000 else NN2 := NMax;
            BeginTrans;
            ExecuteSQL('UPDATE LIGNEOUVPLAT SET ' + SetSQL + ' WHERE BOP_NATUREPIECEG="' + ListeNat[i] + '" AND BOP_SOUCHE="' + ListeSouche[i] + '" AND BOP_NUMERO>=' +
                        IntToStr(NN1) + ' AND BOP_NUMERO<=' + IntToStr(NN2) + ' ' + SetCondition );
            CommitTrans;
          end;
        end;
        // Tranche suivante
        NbDebutTranche := NbDebutTranche + 1000000;
        NbFinTranche := NbFinTranche + 1000000;
      end;
      // sinon trou ds la numéraotation, update pour la souche
    end else ExecuteSQL('UPDATE LIGNEOUVPLAT SET ' + SetSQL + ' WHERE BOP_NATUREPIECEG="' + ListeNat[i] + '" AND BOP_SOUCHE="' + ListeSouche[i] + '"' + ' ' + SetCondition);
  end;
  ListeNat.Clear;
  ListeNat.Free;
  ListeSouche.Clear;
  ListeSouche.Free;
end;

procedure TMajStructBTP.UpDateDecoupeLigne(SetSQL: string; SetCondition : string ='');
var NMin, NMax, NCount, NN1, NN2, NbStep: integer;
  ListeNat, ListeSouche: TStrings;
  Q: TQuery;
  QMax: TQuery;
  i, k: integer;
  NbDebutTranche: integer;
  NbFinTranche: integer;
  //PremierNumero: integer;
  DernierNumero: integer;
  //SQL : string ;
begin
  // Lecture des natures
  ListeNat := TStringList.Create;
  ListeSouche := TStringList.Create;
  Q := OpenSQL('Select DISTINCT GL_NATUREPIECEG, GL_SOUCHE FROM LIGNE', True);
  while not Q.EOF do
  begin
    ListeNat.Add(Q.Fields[0].AsString);
    ListeSouche.Add(Q.Fields[1].AsString);
    Q.Next;
  end;
  Ferme(Q);
  // Balayage des pièces avec découpe
  for i := 0 to ListeNat.Count - 1 do
  begin
    // Récupération du plus grand numéros possible
    QMax := OpenSQL('SELECT MIN(GL_NUMERO), MAX(GL_NUMERO),COUNT(*) FROM LIGNE WHERE GL_NATUREPIECEG="' + ListeNat[i] + '" AND GL_SOUCHE="' + ListeSouche[i] +
                    '"', True);
    //PremierNumero := QMax.Fields[0].AsInteger;
    DernierNumero := QMax.Fields[1].AsInteger;
    NCount := QMax.Fields[2].AsInteger;
    Ferme(QMax);

    if NCount > 50000 then
    begin
      // Première tranche traitée : 0 à 1000000
      NbDebutTranche := 0;
      NbFinTranche := 1000000;

      while (NbDebutTranche < DernierNumero) do
      begin
        // Récupère le premier et le dernier numéro à traiter dans la tranche
        QMax := OpenSQL('SELECT MIN(GL_NUMERO), MAX(GL_NUMERO),COUNT(*) FROM LIGNE WHERE GL_NATUREPIECEG="' + ListeNat[i] + '" AND GL_SOUCHE="' +
          ListeSouche[i] + '" AND GL_NUMERO>=' + IntToStr(NbDebutTranche) + ' AND GL_NUMERO<' + IntToStr(NbFinTranche) + '', True);
        NMin := QMax.Fields[0].AsInteger;
        NMax := QMax.Fields[1].AsInteger;
        NCount := QMax.Fields[2].AsInteger;
        Ferme(QMax);

        if NCount > 0 then
        begin
          NbStep := (NMax - NMin) div 25000;
          for k := 0 to NbStep do
          begin
            NN1 := NMin + k * 25000;
            if k < NbStep then NN2 := NN1 + 25000 else NN2 := NMax;
            BeginTrans;
            ExecuteSQL('UPDATE LIGNE SET ' + SetSQL + ' WHERE GL_NATUREPIECEG="' + ListeNat[i] + '" AND GL_SOUCHE="' + ListeSouche[i] + '" AND GL_NUMERO>=' +
                        IntToStr(NN1) + ' AND GL_NUMERO<=' + IntToStr(NN2) + ' ' + SetCondition );
            CommitTrans;
          end;
        end;
        // Tranche suivante
        NbDebutTranche := NbDebutTranche + 1000000;
        NbFinTranche := NbFinTranche + 1000000;
      end;
      // sinon trou ds la numéraotation, update pour la souche
    end else ExecuteSQL('UPDATE LIGNE SET ' + SetSQL + ' WHERE GL_NATUREPIECEG="' + ListeNat[i] + '" AND GL_SOUCHE="' + ListeSouche[i] + '"' + ' ' + SetCondition);
  end;
  ListeNat.Clear;
  ListeNat.Free;
  ListeSouche.Clear;
  ListeSouche.Free;
end;

procedure TMajStructBTP.UpDateDecoupeLigneBase(SetSQL: string; SetCondition : string );
var NMin, NMax, NCount, NN1, NN2, NbStep: integer;
  ListeNat, ListeSouche: TStrings;
  Q: TQuery;
  QMax: TQuery;
  i, k: integer;
  NbDebutTranche: integer;
  NbFinTranche: integer;
  //PremierNumero: integer;
  DernierNumero: integer;
  //SQL : string ;
begin
  // Lecture des natures
  ListeNat := TStringList.Create;
  ListeSouche := TStringList.Create;
  Q := OpenSQL('Select DISTINCT BLB_NATUREPIECEG, BLB_SOUCHE FROM LIGNEBASE', True);
  while not Q.EOF do
  begin
    ListeNat.Add(Q.Fields[0].AsString);
    ListeSouche.Add(Q.Fields[1].AsString);
    Q.Next;
  end;
  Ferme(Q);
  // Balayage des pièces avec découpe
  for i := 0 to ListeNat.Count - 1 do
  begin
    // Récupération du plus grand numéros possible
    QMax := OpenSQL('SELECT MIN(BLB_NUMERO), MAX(BLB_NUMERO),COUNT(*) FROM LIGNEBASE WHERE BLB_NATUREPIECEG="' + ListeNat[i] + '" AND BLB_SOUCHE="' + ListeSouche[i] +
                    '"', True);
    //PremierNumero := QMax.Fields[0].AsInteger;
    DernierNumero := QMax.Fields[1].AsInteger;
    NCount := QMax.Fields[2].AsInteger;
    Ferme(QMax);

    if NCount > 50000 then
    begin
      // Première tranche traitée : 0 à 1000000
      NbDebutTranche := 0;
      NbFinTranche := 1000000;

      while (NbDebutTranche < DernierNumero) do
      begin
        // Récupère le premier et le dernier numéro à traiter dans la tranche
        QMax := OpenSQL('SELECT MIN(BLB_NUMERO), MAX(BLB_NUMERO),COUNT(*) FROM LIGNEBASE WHERE BLB_NATUREPIECEG="' + ListeNat[i] + '" AND BLB_SOUCHE="' +
          ListeSouche[i] + '" AND BLB_NUMERO>=' + IntToStr(NbDebutTranche) + ' AND BLB_NUMERO<' + IntToStr(NbFinTranche) + '', True);
        NMin := QMax.Fields[0].AsInteger;
        NMax := QMax.Fields[1].AsInteger;
        NCount := QMax.Fields[2].AsInteger;
        Ferme(QMax);

        if NCount > 0 then
        begin
          NbStep := (NMax - NMin) div 25000;
          for k := 0 to NbStep do
          begin
            NN1 := NMin + k * 25000;
            if k < NbStep then NN2 := NN1 + 25000 else NN2 := NMax;
            BeginTrans;
            ExecuteSQL('UPDATE LIGNEBASE SET ' + SetSQL + ' WHERE BLB_NATUREPIECEG="' + ListeNat[i] + '" AND BLB_SOUCHE="' + ListeSouche[i] + '" AND BLB_NUMERO>=' +
                        IntToStr(NN1) + ' AND BLB_NUMERO<=' + IntToStr(NN2) + ' ' + SetCondition );
            CommitTrans;
          end;
        end;
        // Tranche suivante
        NbDebutTranche := NbDebutTranche + 1000000;
        NbFinTranche := NbFinTranche + 1000000;
      end;
      // sinon trou ds la numéraotation, update pour la souche
    end else ExecuteSQL('UPDATE LIGNEBASE SET ' + SetSQL + ' WHERE BLB_NATUREPIECEG="' + ListeNat[i] + '" AND BLB_SOUCHE="' + ListeSouche[i] + '"' + ' ' + SetCondition);
  end;
  ListeNat.Clear;
  ListeNat.Free;
  ListeSouche.Clear;
  ListeSouche.Free;
end;


procedure TMajStructBTP.UpDateDecoupeLigneCompl(SetSQL: string; SetCondition : string ='');
var NMin, NMax, NCount, NN1, NN2, NbStep: integer;
  ListeNat, ListeSouche: TStrings;
  Q: TQuery;
  QMax: TQuery;
  i, k: integer;
  NbDebutTranche: integer;
  NbFinTranche: integer;
  //PremierNumero: integer;
  DernierNumero: integer;
  //SQL : string ;
begin
  // Lecture des natures
  ListeNat := TStringList.Create;
  ListeSouche := TStringList.Create;
  Q := OpenSQL('Select DISTINCT GLC_NATUREPIECEG, GLC_SOUCHE FROM LIGNECOMPL', True);
  while not Q.EOF do
  begin
    ListeNat.Add(Q.Fields[0].AsString);
    ListeSouche.Add(Q.Fields[1].AsString);
    Q.Next;
  end;
  Ferme(Q);
  // Balayage des pièces avec découpe
  for i := 0 to ListeNat.Count - 1 do
  begin
    // Récupération du plus grand numéros possible
    QMax := OpenSQL('SELECT MIN(GLC_NUMERO), MAX(GLC_NUMERO),COUNT(*) FROM LIGNECOMPL WHERE GLC_NATUREPIECEG="' + ListeNat[i] + '" AND GLC_SOUCHE="' + ListeSouche[i] +
                    '"', True);
    //PremierNumero := QMax.Fields[0].AsInteger;
    DernierNumero := QMax.Fields[1].AsInteger;
    NCount := QMax.Fields[2].AsInteger;
    Ferme(QMax);

    if NCount > 50000 then
    begin
      // Première tranche traitée : 0 à 1000000
      NbDebutTranche := 0;
      NbFinTranche := 1000000;

      while (NbDebutTranche < DernierNumero) do
      begin
        // Récupère le premier et le dernier numéro à traiter dans la tranche
        QMax := OpenSQL('SELECT MIN(GLC_NUMERO), MAX(GLC_NUMERO),COUNT(*) FROM LIGNECOMPL WHERE GLC_NATUREPIECEG="' + ListeNat[i] + '" AND GLC_SOUCHE="' +
          ListeSouche[i] + '" AND GLC_NUMERO>=' + IntToStr(NbDebutTranche) + ' AND GLC_NUMERO<' + IntToStr(NbFinTranche) + '', True);
        NMin := QMax.Fields[0].AsInteger;
        NMax := QMax.Fields[1].AsInteger;
        NCount := QMax.Fields[2].AsInteger;
        Ferme(QMax);

        if NCount > 0 then
        begin
          NbStep := (NMax - NMin) div 25000;
          for k := 0 to NbStep do
          begin
            NN1 := NMin + k * 25000;
            if k < NbStep then NN2 := NN1 + 25000 else NN2 := NMax;
            BeginTrans;
            ExecuteSQL('UPDATE LIGNECOMPL SET ' + SetSQL + ' WHERE GLC_NATUREPIECEG="' + ListeNat[i] + '" AND GLC_SOUCHE="' + ListeSouche[i] + '" AND GLC_NUMERO>=' +
                        IntToStr(NN1) + ' AND GLC_NUMERO<=' + IntToStr(NN2) + ' ' + SetCondition );
            CommitTrans;
          end;
        end;
        // Tranche suivante
        NbDebutTranche := NbDebutTranche + 1000000;
        NbFinTranche := NbFinTranche + 1000000;
      end;
      // sinon trou ds la numéraotation, update pour la souche
    end else ExecuteSQL('UPDATE LIGNECOMPL SET ' + SetSQL + ' WHERE GLC_NATUREPIECEG="' + ListeNat[i] + '" AND GLC_SOUCHE="' + ListeSouche[i] + '"' + ' ' + SetCondition);
  end;
  ListeNat.Clear;
  ListeNat.Free;
  ListeSouche.Clear;
  ListeSouche.Free;
end;

procedure TMajStructBTP.UpDateDecoupeLigneFac(SetSQL: string; SetCondition : string ='');
var NMin, NMax, NCount, NN1, NN2, NbStep: integer;
  ListeNat, ListeSouche: TStrings;
  Q: TQuery;
  QMax: TQuery;
  i, k: integer;
  NbDebutTranche: integer;
  NbFinTranche: integer;
  //PremierNumero: integer;
  DernierNumero: integer;
  //SQL : string ;
begin
  // Lecture des natures
  ListeNat := TStringList.Create;
  ListeSouche := TStringList.Create;
  Q := OpenSQL('Select DISTINCT BLF_NATUREPIECEG, BLF_SOUCHE FROM LIGNEFAC', True);
  while not Q.EOF do
  begin
    ListeNat.Add(Q.Fields[0].AsString);
    ListeSouche.Add(Q.Fields[1].AsString);
    Q.Next;
  end;
  Ferme(Q);
  // Balayage des pièces avec découpe
  for i := 0 to ListeNat.Count - 1 do
  begin
    // Récupération du plus grand numéros possible
    QMax := OpenSQL('SELECT MIN(BLF_NUMERO), MAX(BLF_NUMERO),COUNT(*) FROM LIGNEFAC WHERE BLF_NATUREPIECEG="' + ListeNat[i] + '" AND BLF_SOUCHE="' + ListeSouche[i] +
                    '"', True);
    //PremierNumero := QMax.Fields[0].AsInteger;
    DernierNumero := QMax.Fields[1].AsInteger;
    NCount := QMax.Fields[2].AsInteger;
    Ferme(QMax);

    if NCount > 50000 then
    begin
      // Première tranche traitée : 0 à 1000000
      NbDebutTranche := 0;
      NbFinTranche := 1000000;

      while (NbDebutTranche < DernierNumero) do
      begin
        // Récupère le premier et le dernier numéro à traiter dans la tranche
        QMax := OpenSQL('SELECT MIN(BLF_NUMERO), MAX(BLF_NUMERO),COUNT(*) FROM LIGNEFAC WHERE BLF_NATUREPIECEG="' + ListeNat[i] + '" AND BLF_SOUCHE="' +
          ListeSouche[i] + '" AND BLF_NUMERO>=' + IntToStr(NbDebutTranche) + ' AND BLF_NUMERO<' + IntToStr(NbFinTranche) + '', True);
        NMin := QMax.Fields[0].AsInteger;
        NMax := QMax.Fields[1].AsInteger;
        NCount := QMax.Fields[2].AsInteger;
        Ferme(QMax);

        if NCount > 0 then
        begin
          NbStep := (NMax - NMin) div 25000;
          for k := 0 to NbStep do
          begin
            NN1 := NMin + k * 25000;
            if k < NbStep then NN2 := NN1 + 25000 else NN2 := NMax;
            BeginTrans;
            ExecuteSQL('UPDATE LIGNEFAC SET ' + SetSQL + ' WHERE BLF_NATUREPIECEG="' + ListeNat[i] + '" AND BLF_SOUCHE="' + ListeSouche[i] + '" AND BLF_NUMERO>=' +
                        IntToStr(NN1) + ' AND BLF_NUMERO<=' + IntToStr(NN2) + ' ' + SetCondition );
            CommitTrans;
          end;
        end;
        // Tranche suivante
        NbDebutTranche := NbDebutTranche + 1000000;
        NbFinTranche := NbFinTranche + 1000000;
      end;
      // sinon trou ds la numéraotation, update pour la souche
    end else ExecuteSQL('UPDATE LIGNEFAC SET ' + SetSQL + ' WHERE BLF_NATUREPIECEG="' + ListeNat[i] + '" AND BLF_SOUCHE="' + ListeSouche[i] + '"' + ' ' + SetCondition);
  end;
  ListeNat.Clear;
  ListeNat.Free;
  ListeSouche.Clear;
  ListeSouche.Free;
end;

procedure TMajStructBTP.UpDateDecoupeLigneOUV(SetSQL: string; SetCondition : string ='');
var NMin, NMax, NCount, NN1, NN2, NbStep: integer;
  ListeNat, ListeSouche: TStrings;
  Q: TQuery;
  QMax: TQuery;
  i, k: integer;
  NbDebutTranche: integer;
  NbFinTranche: integer;
  //PremierNumero: integer;
  DernierNumero: integer;
  //SQL : string ;
begin
  // Lecture des natures
  ListeNat := TStringList.Create;
  ListeSouche := TStringList.Create;
  Q := OpenSQL('Select DISTINCT BLO_NATUREPIECEG, BLO_SOUCHE FROM LIGNEOUV', True);
  while not Q.EOF do
  begin
    ListeNat.Add(Q.Fields[0].AsString);
    ListeSouche.Add(Q.Fields[1].AsString);
    Q.Next;
  end;
  Ferme(Q);
  // Balayage des pièces avec découpe
  for i := 0 to ListeNat.Count - 1 do
  begin
    // Récupération du plus grand numéros possible
    QMax := OpenSQL('SELECT MIN(BLO_NUMERO), MAX(BLO_NUMERO),COUNT(*) FROM LIGNEOUV WHERE BLO_NATUREPIECEG="' + ListeNat[i] + '" AND BLO_SOUCHE="' + ListeSouche[i] +
                    '"', True);
    //PremierNumero := QMax.Fields[0].AsInteger;
    DernierNumero := QMax.Fields[1].AsInteger;
    NCount := QMax.Fields[2].AsInteger;
    Ferme(QMax);

    if NCount > 50000 then
    begin
      // Première tranche traitée : 0 à 1000000
      NbDebutTranche := 0;
      NbFinTranche := 1000000;

      while (NbDebutTranche < DernierNumero) do
      begin
        // Récupère le premier et le dernier numéro à traiter dans la tranche
        QMax := OpenSQL('SELECT MIN(BLO_NUMERO), MAX(BLO_NUMERO),COUNT(*) FROM LIGNEOUV WHERE BLO_NATUREPIECEG="' + ListeNat[i] + '" AND BLO_SOUCHE="' +
          ListeSouche[i] + '" AND BLO_NUMERO>=' + IntToStr(NbDebutTranche) + ' AND BLO_NUMERO<' + IntToStr(NbFinTranche) + '', True);
        NMin := QMax.Fields[0].AsInteger;
        NMax := QMax.Fields[1].AsInteger;
        NCount := QMax.Fields[2].AsInteger;
        Ferme(QMax);

        if NCount > 0 then
        begin
          NbStep := (NMax - NMin) div 25000;
          for k := 0 to NbStep do
          begin
            NN1 := NMin + k * 25000;
            if k < NbStep then NN2 := NN1 + 25000 else NN2 := NMax;
            BeginTrans;
            ExecuteSQL('UPDATE LIGNEOUV SET ' + SetSQL + ' WHERE BLO_NATUREPIECEG="' + ListeNat[i] + '" AND BLO_SOUCHE="' + ListeSouche[i] + '" AND BLO_NUMERO>=' +
                        IntToStr(NN1) + ' AND BLO_NUMERO<=' + IntToStr(NN2) + ' ' + SetCondition );
            CommitTrans;
          end;
        end;
        // Tranche suivante
        NbDebutTranche := NbDebutTranche + 1000000;
        NbFinTranche := NbFinTranche + 1000000;
      end;
      // sinon trou ds la numéraotation, update pour la souche
    end else ExecuteSQL('UPDATE LIGNEOUV SET ' + SetSQL + ' WHERE BLO_NATUREPIECEG="' + ListeNat[i] + '" AND BLO_SOUCHE="' + ListeSouche[i] + '"' + ' ' + SetCondition);
  end;
  ListeNat.Clear;
  ListeNat.Free;
  ListeSouche.Clear;
  ListeSouche.Free;
end;

procedure TMajStructBTP.UpDateDecoupePiece(SetSQL: string; SetCondition : string ='');
var NMin, NMax, NCount, NN1, NN2, NbStep: integer;
  ListeNat, ListeSouche: TStrings;
  Q: TQuery;
  QMax: TQuery;
  i, k: integer;
  NbDebutTranche: integer;
  NbFinTranche: integer;
  //PremierNumero: integer;
  DernierNumero: integer;
begin
  // Lecture des natures
  ListeNat := TStringList.Create;
  ListeSouche := TStringList.Create;
  Q := OpenSQL('Select DISTINCT GP_NATUREPIECEG, GP_SOUCHE FROM PIECE', True);
  while not Q.EOF do
  begin
    ListeNat.Add(Q.Fields[0].AsString);
    ListeSouche.Add(Q.Fields[1].AsString);
    Q.Next;
  end;
  Ferme(Q);

  // Balayage des pièces avec découpe
  for i := 0 to ListeNat.Count - 1 do
  begin
    // Récupération du plus grand numéros possible
    QMax := OpenSQL('SELECT MIN(GP_NUMERO), MAX(GP_NUMERO),COUNT(*) FROM PIECE WHERE GP_NATUREPIECEG="' + ListeNat[i] + '" AND GP_SOUCHE="' + ListeSouche[i] +
      '"', True);
    //PremierNumero := QMax.Fields[0].AsInteger;
    DernierNumero := QMax.Fields[1].AsInteger;
    NCount := QMax.Fields[2].AsInteger;
    Ferme(QMax);

    // si plus de 50000, update par tranche
    if NCount > 50000 then
    begin
      // Première tranche traitée : 0 à 1000000
      NbDebutTranche := 0;
      NbFinTranche := 1000000;

      while (NbDebutTranche < DernierNumero) do
      begin
        // Récupère le premier et le dernier numéro à traiter dans la tranche
        QMax := OpenSQL('SELECT MIN(GP_NUMERO), MAX(GP_NUMERO),COUNT(*) FROM PIECE WHERE GP_NATUREPIECEG="' + ListeNat[i] + '" AND GP_SOUCHE="' +
          ListeSouche[i] + '" AND GP_NUMERO>=' + IntToStr(NbDebutTranche) + ' AND GP_NUMERO<' + IntToStr(NbFinTranche) + '', True);
        NMin := QMax.Fields[0].AsInteger;
        NMax := QMax.Fields[1].AsInteger;
        NCount := QMax.Fields[2].AsInteger;
        Ferme(QMax);

        if NCount > 0 then
        begin
          NbStep := (NMax - NMin) div 25000;
          for k := 0 to NbStep do
          begin
            NN1 := NMin + k * 25000;
            if k < NbStep then NN2 := NN1 + 25000 else NN2 := NMax;
            BeginTrans;
            ExecuteSQL('UPDATE PIECE SET ' + SetSQL + ' WHERE GP_NATUREPIECEG="' + ListeNat[i] + '" AND GP_SOUCHE="' + ListeSouche[i] + '" AND GP_NUMERO>=' +
              IntToStr(NN1) + ' AND GP_NUMERO<=' + IntToStr(NN2)+ ' ' + SetCondition );
            CommitTrans;
          end;
        end;
        // Tranche suivante
        NbDebutTranche := NbDebutTranche + 1000000;
        NbFinTranche := NbFinTranche + 1000000;
      end;
    end
      // sinon trou ds la numérotation, update pour la souche
    else
    begin
      ExecuteSQL('UPDATE PIECE SET ' + SetSQL
        + ' WHERE GP_NATUREPIECEG="' + ListeNat[i] + '" AND GP_SOUCHE="' + ListeSouche[i] + '"'+ ' ' + SetCondition );
    end;
  end;
  ListeNat.Clear;
  ListeNat.Free;
  ListeSouche.Clear;
  ListeSouche.Free;
end;

procedure TMajStructBTP.UpDateDecoupePiedBase(SetSQL: string; SetCondition : string ='');
var NMin, NMax, NCount, NN1, NN2, NbStep: integer;
  ListeNat, ListeSouche: TStrings;
  Q: TQuery;
  QMax: TQuery;
  i, k: integer;
  NbDebutTranche: integer;
  NbFinTranche: integer;
  //PremierNumero: integer;
  DernierNumero: integer;
begin
  // Lecture des natures
  ListeNat := TStringList.Create;
  ListeSouche := TStringList.Create;
  Q := OpenSQL('Select DISTINCT GPB_NATUREPIECEG, GPB_SOUCHE FROM PIEDBASE', True);
  while not Q.EOF do
  begin
    ListeNat.Add(Q.Fields[0].AsString);
    ListeSouche.Add(Q.Fields[1].AsString);
    Q.Next;
  end;
  Ferme(Q);

  // Balayage des pièces avec découpe
  for i := 0 to ListeNat.Count - 1 do
  begin
    // Récupération du plus grand numéros possible
    QMax := OpenSQL('SELECT MIN(GPB_NUMERO), MAX(GPB_NUMERO),COUNT(*) FROM PIEDBASE WHERE GPB_NATUREPIECEG="' + ListeNat[i] + '" AND GPB_SOUCHE="' + ListeSouche[i] +
      '"', True);
    //PremierNumero := QMax.Fields[0].AsInteger;
    DernierNumero := QMax.Fields[1].AsInteger;
    NCount := QMax.Fields[2].AsInteger;
    Ferme(QMax);

    // si plus de 50000, update par tranche
    if NCount > 50000 then
    begin
      // Première tranche traitée : 0 à 1000000
      NbDebutTranche := 0;
      NbFinTranche := 1000000;

      while (NbDebutTranche < DernierNumero) do
      begin
        // Récupère le premier et le dernier numéro à traiter dans la tranche
        QMax := OpenSQL('SELECT MIN(GPB_NUMERO), MAX(GPB_NUMERO),COUNT(*) FROM PIEDBASE WHERE GPB_NATUREPIECEG="' + ListeNat[i] + '" AND GPB_SOUCHE="' +
          ListeSouche[i] + '" AND GPB_NUMERO>=' + IntToStr(NbDebutTranche) + ' AND GPB_NUMERO<' + IntToStr(NbFinTranche) + '', True);
        NMin := QMax.Fields[0].AsInteger;
        NMax := QMax.Fields[1].AsInteger;
        NCount := QMax.Fields[2].AsInteger;
        Ferme(QMax);

        if NCount > 0 then
        begin
          NbStep := (NMax - NMin) div 25000;
          for k := 0 to NbStep do
          begin
            NN1 := NMin + k * 25000;
            if k < NbStep then NN2 := NN1 + 25000 else NN2 := NMax;
            BeginTrans;
            ExecuteSQL('UPDATE PIEDBASE SET ' + SetSQL + ' WHERE GPB_NATUREPIECEG="' + ListeNat[i] + '" AND GPB_SOUCHE="' + ListeSouche[i] + '" AND GPB_NUMERO>=' +
              IntToStr(NN1) + ' AND GPB_NUMERO<=' + IntToStr(NN2)+ ' ' + SetCondition );
            CommitTrans;
          end;
        end;
        // Tranche suivante
        NbDebutTranche := NbDebutTranche + 1000000;
        NbFinTranche := NbFinTranche + 1000000;
      end;
    end
      // sinon trou ds la numérotation, update pour la souche
    else
    begin
      ExecuteSQL('UPDATE PIEDBASE SET ' + SetSQL
        + ' WHERE GPB_NATUREPIECEG="' + ListeNat[i] + '" AND GPB_SOUCHE="' + ListeSouche[i] + '"'+ ' ' + SetCondition );
    end;
  end;
  ListeNat.Clear;
  ListeNat.Free;
  ListeSouche.Clear;
  ListeSouche.Free;
end;


procedure TMajStructBTP.UpDateDecoupePieceRG(SetSQL: string; SetCondition : string ='');
var NMin, NMax, NCount, NN1, NN2, NbStep: integer;
  ListeNat, ListeSouche: TStrings;
  Q: TQuery;
  QMax: TQuery;
  i, k: integer;
  NbDebutTranche: integer;
  NbFinTranche: integer;
  //PremierNumero: integer;
  DernierNumero: integer;
begin
  // Lecture des natures
  ListeNat := TStringList.Create;
  ListeSouche := TStringList.Create;
  Q := OpenSQL('Select DISTINCT PRG_NATUREPIECEG, PRG_SOUCHE FROM PIECERG', True);
  while not Q.EOF do
  begin
    ListeNat.Add(Q.Fields[0].AsString);
    ListeSouche.Add(Q.Fields[1].AsString);
    Q.Next;
  end;
  Ferme(Q);

  // Balayage des pièces avec découpe
  for i := 0 to ListeNat.Count - 1 do
  begin
    // Récupération du plus grand numéros possible
    QMax := OpenSQL('SELECT MIN(PRG_NUMERO), MAX(PRG_NUMERO),COUNT(*) FROM PIECERG WHERE PRG_NATUREPIECEG="' + ListeNat[i] + '" AND PRG_SOUCHE="' + ListeSouche[i] +
      '"', True);
    //PremierNumero := QMax.Fields[0].AsInteger;
    DernierNumero := QMax.Fields[1].AsInteger;
    NCount := QMax.Fields[2].AsInteger;
    Ferme(QMax);

    // si plus de 50000, update par tranche
    if NCount > 50000 then
    begin
      // Première tranche traitée : 0 à 1000000
      NbDebutTranche := 0;
      NbFinTranche := 1000000;

      while (NbDebutTranche < DernierNumero) do
      begin
        // Récupère le premier et le dernier numéro à traiter dans la tranche
        QMax := OpenSQL('SELECT MIN(PRG_NUMERO), MAX(PRG_NUMERO),COUNT(*) FROM PIECERG WHERE PRG_NATUREPIECEG="' + ListeNat[i] + '" AND PRG_SOUCHE="' +
          ListeSouche[i] + '" AND PRG_NUMERO>=' + IntToStr(NbDebutTranche) + ' AND PRG_NUMERO<' + IntToStr(NbFinTranche) + '', True);
        NMin := QMax.Fields[0].AsInteger;
        NMax := QMax.Fields[1].AsInteger;
        NCount := QMax.Fields[2].AsInteger;
        Ferme(QMax);

        if NCount > 0 then
        begin
          NbStep := (NMax - NMin) div 25000;
          for k := 0 to NbStep do
          begin
            NN1 := NMin + k * 25000;
            if k < NbStep then NN2 := NN1 + 25000 else NN2 := NMax;
            BeginTrans;
            ExecuteSQL('UPDATE PIECERG SET ' + SetSQL + ' WHERE PRG_NATUREPIECEG="' + ListeNat[i] + '" AND PRG_SOUCHE="' + ListeSouche[i] + '" AND PRG_NUMERO>=' +
              IntToStr(NN1) + ' AND PRG_NUMERO<=' + IntToStr(NN2)+ ' ' + SetCondition );
            CommitTrans;
          end;
        end;
        // Tranche suivante
        NbDebutTranche := NbDebutTranche + 1000000;
        NbFinTranche := NbFinTranche + 1000000;
      end;
    end
      // sinon trou ds la numérotation, update pour la souche
    else
    begin
      ExecuteSQL('UPDATE PIECERG SET ' + SetSQL
        + ' WHERE PRG_NATUREPIECEG="' + ListeNat[i] + '" AND PRG_SOUCHE="' + ListeSouche[i] + '"'+ ' ' + SetCondition );
    end;
  end;
  ListeNat.Clear;
  ListeNat.Free;
  ListeSouche.Clear;
  ListeSouche.Free;
end;

procedure TMajStructBTP.TraiteModifVue(TheListRef: TOB);
var TOBREF,TOBDest : TOB;
begin

  TOBREF := GetVueRef(TheListRef.getValue('BTV_NOMELT'));
  if TOBREF = nil then exit;

  TOBDest := GetVueDestination(TheListRef.getValue('BTV_NOMELT'));

  if not IsTraiteVue (TheListRef,TOBRef,TOBDest) then
  begin
    TOBREf.free;
    if TOBdest <> nil then TOBDest.free;
    exit; // --> pas a traiter
  end;

  TOBREF.SetAllModifie(True);
  TOBREF.InsertOrUpdateDB;
  LanceMajVues := True;
  TOBREf.free;
  TOBDest.free;
end;

function TMajStructBTP.GetVueRef (NomVue : string) : TOB;
var QQ : TQuery;
    TOBResult : TOB;
begin
  result := nil;
  QQ := OpenSQLDb (DBRef,'SELECT * FROM DEVUES WHERE DV_NOMVUE="'+NomVue+'"',true);
  //if QQ.eof then BEGIN fCodeStatus := 120; Exit; END;
  if QQ.eof then BEGIN Exit; END;
  TOBresult := TOB.Create ('DEVUES',nil,-1);
  TOBresult.SelectDB ('',QQ);
  ferme (QQ);
  result := TOBresult;
end;

function TMajStructBTP.GetVueDestination(NomVue: string): TOB;
var QQ : TQuery;
    TOBResult : TOB;
begin
  result := nil;
  QQ := OpenSQL ('SELECT * FROM DEVUES WHERE DV_NOMVUE="'+NomVue+'"',true);
  if QQ.eof then BEGIN Exit; END;
  TOBresult := TOB.Create ('DEVUES',nil,-1);
  TOBresult.SelectDB ('',QQ);
  ferme (QQ);
  result := TOBresult;
end;

function TMajStructBTP.IsTraiteVue (TheListRef,TOBRef,TOBDest : TOB) : boolean;
begin
  result := true;
  if TOBDest = nil then exit; // -> pas de destination c'est donc une création
  if (ForceGlobale) or (IsModeMajHalley) then Exit;
  if (TheListRef.GetString('BTV_FORCE') = 'X') then exit; // --> Force donc on lance
  if TOBDest.getValue('DV_NUMVERSION') < TOBRef.getValue('DV_NUMVERSION') then exit; // version sur ref > cele de la base
  result := false;
end;

procedure TMajStructBTP.DropVuesPhysiques;
begin
	XX.Titre.Caption := 'Suppression des vues physiques';
  XX.Titre.visible := true;
  XX.Titre.Refresh;
  XX.PB.Visible := true;
  XX.Refresh;
  XX.CURVUE.visible := true;
  DBDeleteAllView (DBSOC,XX.CURVUE,DestDriver);
  XX.CURVUE.visible := false;
end;

procedure TMajStructBTP.CreateVuesPhysiques;
begin
	XX.Titre.Caption := 'Création des vues physiques';
  XX.Titre.visible := true;
  XX.Titre.Refresh;
  XX.PB.Visible := true;
  XX.CURVUE.visible := true;
  XX.Refresh;
  DBCreateAllView (DBSOC,XX.CURVUE,DestDriver);
  XX.CURVUE.visible := false;
end;

procedure TMajStructBTP.TraiteChangementMenu(TheListRef: TOB);
var NumeroMenu : string;
    TheMenusRef,TheMenusDest,TheMenusSpecif : TOB;
    TheMenuRef,TheMenuDest : TOB;
    indice : integer;
begin
  TheMenusref := TOB.create ('LES MENUS SOC',nil,-1);
  TheMenusDest := TOB.create ('LES MENUS SOCD',nil,-1);
  TheMenusSpecif := TOB.create ('LES MENUS SPECIFS',nil,-1);

  NumeroMenu := TheLIstRef.GetValue ('BTV_NOMELT');
	XX.Titre.Caption := 'Mise à jour des Menus';
  XX.Titre.visible := true;
  XX.Titre.Refresh;
  XX.PB.Visible := true;
  XX.Refresh;
  GetMenus (NumeroMenu,TheMenusRef,TheMenusDest);
  GetMenusSpecif (NumeroMenu,TheMenusSpecif);
  // phase de préparation
  for indice := 0 to TheMenusref.detail.count -1 do
  begin
    TheMenuRef :=TheMenusref.detail[Indice];
    TheMenuDest := TheMenusDest.findFirst(['MN_1','MN_TAG'],[TheMenuRef.GetValue('MN_1'),TheMenuRef.GetValue('MN_TAG')],true);
    if TheMenuDest <> nil then
    begin
    	if TheListRef.GetValue ('BTV_WITHDATA') <> 'X' Then
      begin
        // on récupère la confidentialité du dossier client
        TheMenuRef.putValue('MN_ACCESGRP',TheMenuDest.GetValue('MN_ACCESGRP'));
      end;
    end else
    begin
      // dans le cas ou nouveau menu on positionne celui ci comme non visible si la marque dev est en place.
      if TheMenuRef.GetString('MN_VERSIONDEV')='X' then
      begin
      	TheMenuRef.SetString('MN_ACCESGRP',StringOfChar('-',100));
      	TheMenuRef.SetString('MN_VERSIONDEV','-');
      end;
    end;
  end;
  //
  TheMenusref.detail.sort ('MN_1;MN_2;MN_3;MN_4'); // pour pouvoir reintégrer les menus spécifs
  //
  AddmenusSpecif (TheMenusRef,TheMenusSpecif);
  //
  // Phase de nettoyage
  //
  DropMenuDest(NumeroMenu);
  //
  TheMenusDest.clearDetail;
  for indice := 0 to TheMenusref.detail.count -1 do
  begin
    TheMenuRef :=TheMenusref.detail[Indice];
    TheMenuDest := TOB.Create ('MENU',TheMenusDest,-1);
    TheMenuDest.Dupliquer(TheMenuref,false,true);
  end;
  //
  // Phase d'insertion des données
  //
  TheMenusDest.SetAllModifie(true);
  TheMenusDest.InsertDB(nil);
  TheMenusref.free;
  TheMenusDest.free;
  TheMenusSpecif.free;
  //
end;



procedure TMajStructBTP.DropFunction(NomF:String);
var Sql : string;
begin
  Sql := 'SELECT  1 '+
         'FROM    Information_schema.Routines '+
         'WHERE   Specific_schema = "dbo" AND '+
         'specific_name = "'+NomF+'" AND '+
         'Routine_Type = "FUNCTION"';
  if ExisteSQL(Sql) then
  begin
    ExecuteSql('DROP FUNCTION DBO.'+NomF);
  end;
end;

procedure TMajStructBTP.CreateFunctionDecoupage;
var Sql : string;
begin
  Sql := 'CREATE FUNCTION [dbo].[ReadToken] '+
         '( '+
         '@Chaine VarChar(255), '+
         '@Idx TinyInt,	'+
         '@Separateur VarChar(5) = ";" '+
         ')'+
         'RETURNS VarChar(255) '+
         'AS '+
         'BEGIN '+
         'Declare @Valeur VarChar(255) '+
         'Declare @Pos Int '+
         'Declare @Cpt Int '+
         'SELECT @Cpt = 1, @Pos = 0, @Valeur=NULL '+
         'While (@Cpt < @Idx AND Charindex(@Separateur, @Chaine, @Pos+1)>@Pos) '+
         'Begin '+
         'SELECT @Cpt=@Cpt+1, @Pos=Charindex(@Separateur, @Chaine, @Pos+1) '+
         'End '+
         'IF @Cpt=@Idx '+
         'SET @Pos = @Pos + 1 '+
         'IF Charindex(@Separateur, @Chaine, @Pos)>@Pos '+
         'SET @Valeur = ltrim(rtrim(Substring(@Chaine, @Pos, Charindex(@Separateur, @Chaine, @Pos)-@Pos))) '+
         'ELSE IF Charindex(@Separateur, @Chaine, @Pos)=@Pos '+
         'SET @Valeur = "" '+
         'ELSE '+
         'SET @Valeur = ltrim(rtrim(Substring(@Chaine, @Pos, 255))) '+
         'RETURN @Valeur '+
         'END ';
  ExecuteSQL(Sql);
end;


procedure TMajStructBTP.CreateSqlFunctions;
begin
  DropFunction('ReadToken');
  //
  CreateFunctionDecoupage;
end;

procedure TMajStructBTP.CreateRealTable (NomTable : string);
var TheStructureRef : TOB;
begin
  TheStructureRef := TOB.Create ('DETABLES',nil,-1);
  //
  TRY
    if TableExisteDB (NomTable,DBREF) then
    begin
      fCodeStatus := 0;
      GetStructureTableFromRef (TheSTructureRef,NomTable);
      if fCodeStatus <> 0 then exit;
      if not TableExiste (NomTable) then
      begin
        CreeTable (TheStructureRef,NomTable);
        if fCodeStatus <> 0 then exit;
      end;
    end;
    fcodeStatus := 0;
    //
  FINALLY
    TheStructureRef.free;
  END;

end;

procedure TMajStructBTP.ConstitueTableSoucheBTP;
  procedure RecupDatasFromSouche;
  var SQl : String;
  begin
    if ISColumnExists('SH_NUMMOISPIECE','SOUCHE') then
    begin
    SQL := 'INSERT INTO BSOUCHE '+
           '(BS0_ENTITY,BS0_TYPE,BS0_SOUCHE,BS0_NUMMOISPIECE) '+
           'SELECT SH_ENTITY,SH_TYPE,SH_SOUCHE,SH_NUMMOISPIECE FROM SOUCHE';
    ExecuteSQLContOnExcept(SQL);
  end;
  end;

var NomTable : string;
begin
  //
  InitMoveProgressForm(nil, 'Creation de la table BSOUCHE', 'Veuillez patienter SVP ...', 2, FALSE, TRUE);
  MoveCurProgressForm('Constitution de la table Physique ');

  CreateRealTable ('BSOUCHE');
  if fCodeStatus = 0 then
  begin
    MoveCurProgressForm('Récupération des données de SOUCHE');
    RecupDatasFromSouche;
    LanceMajVues := True;
  end;
  FiniMoveProgressForm();

end;

procedure TMajStructBTP.ConstitueTableFonctionBTP;
  procedure RecupDatasFromFonctions;
  var SQl : String;
  begin
    if ISColumnExists('AFO_GEREPLANNING','FONCTION') then
    begin
    SQL := 'INSERT INTO BFONCTION '+
           '(BFO_FONCTION,BFO_BCOLOR,BFO_COLOR,BFO_GEREPLANNING) '+
           'SELECT AFO_FONCTION,AFO_BCOLOR,AFO_COLOR,AFO_GEREPLANNING FROM FONCTION';
    ExecuteSQLContOnExcept(SQL);
  end;
  end;

var NomTable : string;
begin
  InitMoveProgressForm(nil, 'Creation de la table BFONCTION', 'Veuillez patienter SVP ...', 2, FALSE, TRUE);
  MoveCurProgressForm('Constitution de la table Physique ');
  //
  CreateRealTable ('BFONCTION');
  if fCodeStatus = 0 then
  begin
    MoveCurProgressForm('Récupération des données de FONCTION');
    RecupDatasFromFonctions;
    LanceMajVues := True;
  end;
  FiniMoveProgressForm();
end;

procedure TMajStructBTP.ConstitueTableRessourceBTP;
  procedure RecupDatasFromRessource;
  var SQl : String;
  begin
    if ISColumnExists('ARS_GERESAV','RESSOURCE') then
    begin
    SQL := 'INSERT INTO BRESSOURCE '+
           '(BRS_RESSOURCE,BRS_SECTION,BRS_GERESAV) '+
           'SELECT ARS_RESSOURCE,ARS_SECTION,ARS_GERESAV FROM RESSOURCE';
    ExecuteSQLContOnExcept(SQL);
  end;
  end;

var NomTable : string;
begin
  //
  InitMoveProgressForm(nil, 'Creation de la table BRESSOURCE', 'Veuillez patienter SVP ...', 2, FALSE, TRUE);
  MoveCurProgressForm('Constitution de la table Physique ');
  CreateRealTable ('BRESSOURCE');
  if fCodeStatus = 0 then
  begin
    MoveCurProgressForm('Récupération des données de RESSOURCE');
    RecupDatasFromRessource;
    LanceMajVues := True;
  end;
  FiniMoveProgressForm();
end;

procedure TMajStructBTP.ConstitueTableContactBTP;
  procedure RecupDatasFromContacts;
  var SQl : String;
  begin
    //
    if ISColumnExists('C_BATIMENT','CONTACT') then
    begin
    SQL := 'INSERT INTO BCONTACT '+
           '(BC1_TYPECONTACT,BC1_AUXILIAIRE,BC1_NUMEROCONTACT,BC1_NUMEROADRESSE,BC1_BATIMENT,BC1_ETAGE,BC1_ESCALIER,BC1_PORTE) '+
           'SELECT C_TYPECONTACT,C_AUXILIAIRE,C_NUMEROCONTACT,C_NUMEROADRESSE,C_BATIMENT,C_ETAGE,C_ESCALIER,C_PORTE FROM CONTACT';
    end;
    ExecuteSQLContOnExcept(SQL);
  end;

var NomTable : string;
begin
  InitMoveProgressForm(nil, 'Creation de la table BCONTACT', 'Veuillez patienter SVP ...', 2, FALSE, TRUE);
  MoveCurProgressForm('Constitution de la table Physique ');
  //
  CreateRealTable ('BCONTACT');
  if fCodeStatus = 0 then
  begin
    MoveCurProgressForm('Récupération des données de CONTACTS');
    RecupDatasFromContacts;
    LanceMajVues := True;
  end;
  FiniMoveProgressForm();
end;


procedure TMajStructBTP.ConstitueTableTiersBTP;
  procedure RecupDatasFromTIERS;
  var SQl : String;
  begin
    if ISColumnExists ('T_CODEBARRE','TIERS') and ISColumnExists ('T_QUALIFCODEBARRE','TIERS') then
    begin
    SQL := 'INSERT INTO BTIERS '+
           '(BT1_AUXILIAIRE,BT1_IDSPIGAO,BT1_CODEBARRE,BT1_QUALIFCODEB) '+
           'SELECT T_AUXILIAIRE,T_IDSPIGAO,T_CODEBARRE,T_QUALIFCODEBARRE FROM TIERS';
    ExecuteSQLContOnExcept(SQL);
  end;
  end;

var NomTable : string;
begin
  //
  InitMoveProgressForm(nil, 'Creation de la table BTIERS', 'Veuillez patienter SVP ...', 2, FALSE, TRUE);
  MoveCurProgressForm('Constitution de la table Physique ');
  CreateRealTable ('BTIERS');
  if fCodeStatus = 0 then
  begin
    MoveCurProgressForm('Récupération des données de TIERS');
    RecupDatasFromTIERS;
    LanceMajVues := True;
  end;
  FiniMoveProgressForm();
end;

procedure TMajStructBTP.ConstitueTableAdressesBTP;
  procedure RecupDatasFromAdresses;
  var SQl : String;
  begin
    if ISColumnExists('ADR_CODEENTREE2','ADRESSES') then
    begin
    SQL := 'INSERT INTO BADRESSES '+
           '(BA0_TYPEADRESSE,BA0_NUMEROADRESSE,BA0_GARDIEN,BA0_CODEENTREE1,BA0_CODEENTREE2,BA0_CLETELEPHONE,BA0_INT) '+
           'SELECT ADR_TYPEADRESSE,ADR_NUMEROADRESSE,ADR_GARDIEN,ADR_CODEENTREE1,ADR_CODEENTREE2,ADR_CLETELEPHONE,ADR_INT FROM ADRESSES';
    ExecuteSQLContOnExcept(SQL);
  end;
  end;

var NomTable : string;
begin
  //
  InitMoveProgressForm(nil, 'Creation de la table BADRESSES', 'Veuillez patienter SVP ...', 2, FALSE, TRUE);
  MoveCurProgressForm('Constitution de la table Physique ');
  CreateRealTable ('BADRESSES');
  if fCodeStatus = 0 then
  begin
    MoveCurProgressForm('Récupération des données de ADRESSES');
    RecupDatasFromAdresses;
    LanceMajVues := True;
  end;
  FiniMoveProgressForm();
end;


procedure TMajStructBTP.beforeTraitement;
begin
  // Création des fonctions locales a la base de donnée
  CreateSqlFunctions;
  //
	ExecuteSQL('DELETE FROM DECOMBOS WHERE DO_COMBO="WNATURETRAVAIL"');
  if VersionBaseDest < '998.O' then
  begin
    ExecuteSql('DELETE FROM LISTE WHERE LI_LISTE = "IMMULIMMO"');
		ExecuteSql('UPDATE DECHAMPS SET DH_CONTROLE="" WHERE DH_NOMCHAMP LIKE "IRF_TVTS_%"');
  end;
  if VersionBaseDest < '998.Q' then
  begin
    ExecuteSql('DELETE FROM PUBLICOTIS WHERE PUO_PREDEFINI="CEG" AND PUO_TYPEAFFECT="MSA" '+
               'AND PUO_RUBRIQUE="41" AND PUO_UTILSEGMENT="019"');
		ExecuteSql('DELETE FROM INSTITUTIONPAYE WHERE PIP_PREDEFINI = "CEG"');
		ExecuteSql('DELETE FROM PUBLICOTIS WHERE PUO_UTILSEGMENT = "786" AND '+
    					 'PUO_PREDEFINI = "CEG" AND PUO_RUBRIQUE  = "0002"');
    AGLNettoieListes ('PGLANCEBULLAVEC', 'PPU_BULLCONTRAT');
    AGLNettoieListes ('PGDECLARATIONAED', 'PDN_NBANNULEREMP');
    AglNettoieListes ('PGDECLARAEDGEN', 'PDN_NBANNULEREMP');
    AglNettoieListes ('PGENVOISOCIAL', 'PES_REFDECLREMP');
  end;
  if VersionBaseDest < '998.Z4' then
  begin
    ExecuteSql('DELETE FROM PUBLICOTIS WHERE PUO_PREDEFINI="CEG" AND PUO_TYPEAFFECT="MSA" AND PUO_RUBRIQUE="Z33" AND PUO_UTILSEGMENT="019"');
    ExecuteSql ('DELETE FROM LISTE WHERE LI_LISTE="PGEMULMVTABS"');
    ExecuteSql ('DELETE FROM LISTE WHERE LI_LISTE="PGMULMVTABS"');
    ExecuteSql ('DELETE FROM LISTE WHERE LI_LISTE="PGEMULMVTABSR"');
    ExecuteSql ('DELETE FROM LISTE WHERE LI_LISTE="PGEMULMVTABSASS"');
    ExecuteSql ('DELETE FROM LISTE WHERE LI_LISTE="PGSTAGEFORM"');
    ExecuteSql ('DELETE FROM LISTE WHERE LI_LISTE="PGSESSIONSTAGE"');
    ExecuteSql ('DELETE FROM LISTE WHERE LI_LISTE="PGSESSIONSTAGEMS"');
    ExecuteSql ('DELETE FROM LISTE WHERE LI_LISTE="PGCOMPLSALARIEMSA"');
    ExecuteSql ('DELETE FROM LISTE WHERE LI_LISTE="PGERECAPSAL"');
  end;
  if VersionBaseDest < '998.ZF' then
  begin
		ExecuteSql('DELETE FROM LISTE WHERE LI_LISTE ="CPSDDMULGENMANDAT"');
  end;
  //
  if VersionBaseDest < '998.ZZZO' then
  begin
    ExecuteSQL('DELETE FROM COMMUN WHERE CO_TYPE="BS0"');
    ExecuteSQL('DELETE FROM COMMUN WHERE CO_TYPE="BS1"');
    ExecuteSQL('DELETE FROM COMMUN WHERE CO_TYPE="BS2"');
    ExecuteSQL('DELETE FROM COMMUN WHERE CO_TYPE="BAA"');
    ExecuteSQL('DELETE FROM COMMUN WHERE CO_TYPE="BTS"');
    ExecuteSQL('DELETE FROM COMMUN WHERE CO_TYPE="BM7"');
    ExecuteSQL('DELETE FROM COMMUN WHERE CO_TYPE="BM6"');
    if TableExiste ('BTYPELIGBAST') then ExecuteSQL('DELETE FROM BTYPELIGBAST');
  end;
  //
  if VersionBaseDest < '998.ZZZK' then
  begin
    DropVuePourrie('BTPREVFACTPREPA');
    ExecuteSQL('DELETE FROM DEVUES WHERE DV_NOMVUE="BTPREVFACTPREPA"');
  end;
  //
  if not TableExiste('BADRESSES') then
  begin
    ConstitueTableAdressesBTP;
  end;
  //
  if not TableExiste('BTIERS') then
  begin
    ConstitueTableTiersBTP;
  end;
  //
  if not TableExiste('BCONTACT') then
  begin
    ConstitueTableContactBTP;
  end;
  //
  if not TableExiste('BFONCTION') then
  begin
    ConstitueTableFonctionBTP;
  end;
  //
  if not TableExiste('BRESSOURCE') then
  begin
    ConstitueTableRessourceBTP;
  end;

  if not TableExiste('BSOUCHE') then
  begin
    ConstitueTableSoucheBTP;
  end;

  ExecuteSQL('UPDATE MENU SET MN_VERSIONDEV="-" WHERE MN_VERSIONDEV="X"');
end;

procedure TMajStructBTP.AfterTraitement;
var CaisseCp : string;
begin
  // Edition 7 compta
  if VersionBaseDest < '998.E' then
  begin
    SupprimeEtat('E', 'I58', '054', True);
    SupprimeEtat('E', 'I58', '055', True);
    SupprimeEtat('E', 'I58', 'A54', True);
    SupprimeEtat('E', 'I58', 'A55', True);
    SupprimeEtat('E', 'I58', '855', True);
    SupprimeEtat('E', 'I58', '954', True);
    SupprimeEtat('E', 'I58', '955', True);
    SupprimeEtat('E', 'I90', '054', True);
    SupprimeEtat('E', 'I90', '055', True);
    SupprimeEtat('E', 'I90', 'A54', True);
    SupprimeEtat('E', 'I90', 'A55', True);
    // Edition 6 paie
    ExecuteSql('DELETE FROM HISTOSAISRUB WHERE PSD_RUBRIQUE="" AND PSD_ORIGINEMVT="SRB"');
    ExecuteSql('DELETE FROM CONTRATTRAVAIL WHERE PCI_SALARIE NOT IN (SELECT PSA_SALARIE FROM SALARIES)');
    ExecuteSql('DELETE FROM CONTRATTRAVAIL WHERE PCI_SALARIE NOT IN (SELECT PSA_SALARIE FROM SALARIES)');
    ExecuteSql('DELETE FROM PUBLICOTIS WHERE PUO_NATURERUB="REM" AND PUO_RUBRIQUE="09"');
	end ;
  if VersionBaseDest < '998.O' then
  begin
    // edition 7 sp2 compta
    SupprimeEtat('E', 'I58', '054', True);
    SupprimeEtat('E', 'I58', '055', True);
    SupprimeEtat('E', 'I58', 'A54', True);
    SupprimeEtat('E', 'I58', 'A55', True);
    SupprimeEtat('E', 'I58', '855', True);
    SupprimeEtat('E', 'I58', '954', True);
    SupprimeEtat('E', 'I58', '955', True);
    SupprimeEtat('E', 'IK6', '054', True);
    SupprimeEtat('E', 'IK6', '055', True);
    SupprimeEtat('E', 'IK6', 'A54', True);
    SupprimeEtat('E', 'IK6', 'A55', True);
    SupprimeEtat('E', 'IK6', '954', True);
    SupprimeEtat('E', 'IK6', '955', True);
    SupprimeEtat('E', 'I90', '054', True);
    SupprimeEtat('E', 'I90', '055', True);
    SupprimeEtat('E', 'I90', 'A54', True);
    SupprimeEtat('E', 'I90', 'A55', True);
    SupprimeEtat('E', 'I90', '954', True);
    SupprimeEtat('E', 'I90', '955', True);
    ExecuteSql('DELETE FROM IMOTVTS WHERE NOT EXISTS(SELECT IRF_NUM FROM IMOREF WHERE IRF_NUM = ITS_NUM)');
  end;
  if VersionBaseDest < '998.Q' then
  begin
		ExecuteSql('DELETE FROM YEDMFILES WHERE YEF_CUSTOMTYPE LIKE "PG%" AND '+
    					 'YEF_ID IN (SELECT YEM_ID FROM YEDMEXTMETIER WHERE YEM_PREDEFINI = "CEG")');
    ExecuteSql('DELETE FROM YEDMEXTMETIER WHERE YEM_PREDEFINI = "CEG" AND '+
    					 'YEM_ID NOT IN (SELECT YEF_ID FROM YEDMFILES)');
    { Mise à jour des nouveaux champs liés à la caisse C.P. puis suppression de l'ancien paramètre société }
    if ExisteSQL ('SELECT SOC_NOM FROM PARAMSOC WHERE SOC_NOM="SO_PGCAISSECP"') then
    begin
  		 CaisseCP:= GetParamSocSecur('SO_PGCAISSECP', '000');
       if (CaisseCP <> '000') and (CaisseCP <> '') then SetParamSoc ('SO_PGADHESIONCAISSECP', True);
       ExecuteSQL('UPDATE ETABCOMPL SET ETB_NATCAISSECP="'+CaisseCP+'" WHERE ETB_NATCAISSECP IS NULL');
       ExecuteSQL('DELETE FROM PARAMSOC WHERE SOC_NOM="SO_PGCAISSECP"');
    end;
    { Maj des données Concept Contrats de travail }
    ExecuteSQL('UPDATE MENU SET MN_ACCESGRP =(select MN_ACCESGRP  from menu MN where  MN.mn_tag=200062) where mn_1=200 and mn_2=2');
    ExecuteSQL('UPDATE MENU SET MN_ACCESGRP =(select MN_ACCESGRP  from menu MN where  MN.mn_tag=200001) where mn_1=200 and mn_2=1 and mn_3=2');
    ExecuteSQL('UPDATE MENU set mn_libelle=mn_libelle||" Inactif" where mn_tag=200062 and mn_libelle not like "%Inactif%"');
    { Maj du libellé Module Compétences }
    ExecuteSQL('Update menu set mn_libelle = "Compétences" where mn_2=377');
		PgInitTabletteTypeDesti;
  end;
  if VersionBaseDest < '998.Z4' then
  begin
    ExecuteSQL ('UPDATE PAIEPARIM SET PAI_USINGMAINTIEN="-"');
    ExecuteSQL ('UPDATE PAIEPARIM SET PAI_USINGMAINTIEN="X" WHERE (PAI_IDENT=8 OR PAI_IDENT=38 OR PAI_IDENT=45 '+
                'OR PAI_IDENT=46 OR PAI_IDENT=48 OR PAI_IDENT=49 OR PAI_IDENT=60 OR PAI_IDENT=100 '+
                'OR (PAI_IDENT>=157 AND  PAI_IDENT<=161) OR PAI_IDENT=220 OR PAI_IDENT=221 '+
                'OR (PAI_IDENT>=345 AND   PAI_IDENT<=348) OR (PAI_IDENT>=400 AND  PAI_IDENT<=402) '+
                'OR PAI_IDENT=500)');
    //
    ExecuteSQL ('UPDATE SALARIESCOMPL SET PSZ_TYPVARAJUJOUR="ETB" WHERE (PSZ_TYPVARAJUJOUR="" OR PSZ_TYPVARAJUJOUR IS NULL)');
    ExecuteSQL ('UPDATE SALARIESCOMPL SET PSZ_TYPVARAJUHEURE="ETB" WHERE (PSZ_TYPVARAJUHEURE="" OR PSZ_TYPVARAJUHEURE IS NULL)');
    ExecuteSQL ('UPDATE SALARIESCOMPL SET PSZ_TYPVARAJUBASE="ETB" WHERE (PSZ_TYPVARAJUBASE="" OR PSZ_TYPVARAJUBASE IS NULL)');
    ExecuteSQL ('UPDATE SALARIESCOMPL SET PSZ_TYPVARAJUMOIS="ETB" WHERE (PSZ_TYPVARAJUMOIS="" OR PSZ_TYPVARAJUMOIS IS NULL)');
    ExecuteSQL ('UPDATE SALARIESCOMPL SET PSZ_VARAJUJOUR="" WHERE PSZ_VARAJUJOUR IS NULL');
    ExecuteSQL ('UPDATE SALARIESCOMPL SET PSZ_VARAJUHEURE="" WHERE PSZ_VARAJUHEURE IS NULL');
    ExecuteSQL ('UPDATE SALARIESCOMPL SET PSZ_VARAJUBASE="" WHERE PSZ_VARAJUBASE IS NULL');
    ExecuteSQL ('UPDATE SALARIESCOMPL SET PSZ_VARAJUMOIS="" WHERE PSZ_VARAJUMOIS IS NULL');
    ExecuteSQL ('UPDATE ETABCOMPL  SET ETB_VARAJUJOUR="" WHERE ETB_VARAJUJOUR IS NULL');
    ExecuteSQL ('UPDATE ETABCOMPL  SET ETB_VARAJUHEURE="" WHERE ETB_VARAJUHEURE IS NULL');
    ExecuteSQL ('UPDATE ETABCOMPL  SET ETB_VARAJUBASE="" WHERE ETB_VARAJUBASE IS NULL');
    ExecuteSQL ('UPDATE ETABCOMPL  SET ETB_VARAJUMOIS="" WHERE ETB_VARAJUMOIS IS NULL');
    //
    ExecuteSQL ('UPDATE PAIEPARIM SET PAI_USINGMAINTIEN="-"');
    ExecuteSQL ('UPDATE PAIEPARIM SET PAI_USINGMAINTIEN="X" WHERE (PAI_IDENT=8 OR PAI_IDENT=38 OR PAI_IDENT=45 '+
                'OR PAI_IDENT=46 OR PAI_IDENT=48 OR PAI_IDENT=49 OR PAI_IDENT=60 OR PAI_IDENT=100 '+
                'OR (PAI_IDENT>=157 AND  PAI_IDENT<=161) OR PAI_IDENT=220 OR PAI_IDENT=221 '+
                'OR (PAI_IDENT>=345 AND   PAI_IDENT<=348) OR (PAI_IDENT>=400 AND  PAI_IDENT<=402) '+
                'OR PAI_IDENT=500)');
  end;
  if VersionBaseDest < '998.ZE' then // edition 12 et 13  paie et compta
  begin
    // Compta V9 ed 12 et 13
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="CPSDDMULGENMANDAT"');
		ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="CPENCTOUS"');
		ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="CPSDDMULIMPAYES"');
		ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="CPSDDMULHISTO"');
		ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="CPSDDMULVISUEMAND"');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="MULMMVTS"');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="CPVISECRU"');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="CPVISECRS"');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="CPVISECRR"');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="CPVISECRP"');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="CPVISECRN"');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="CPVISECRMU"');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="CPVISECRMS"');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="CPVISECRMR"');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="CPVISECRMP"');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="CPVISECRM"');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="CPVISECRH"');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="CPSUPECRU"');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="CPSUPECRS"');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="CPSUPECRR"');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="CPSUPECRP"');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="CPSUPECRN"');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="CPSUPECRI"');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="CPRECOPECRG"');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="CPMODECRU"');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="CPMODECRS"');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="CPMODECRR"');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="CPMODECRP"');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="CPMODECRN"');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="CPMODECRMU"');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="CPMODECRMS"');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="CPMODECRMR"');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="CPMODECRMP"');
    ExecuteSQL('DELETE FROM LISTE WHERE LI_LISTE ="CPMODECRM"');
    SupprimeEtat('E', 'CSD', 'CSD');
    SupprimeEtat('E', 'CSH', 'CSH');
    SupprimeEtat('E', 'CSI', 'CSI');
    SupprimeEtat('E', 'CSJ', 'CSJ');
    // PAIE V9 Ed 13
    ExecuteSQL ('UPDATE COTISATION SET PCT_TRANCHETAXE="" WHERE PCT_TRANCHETAXE IS NULL');
    ExecuteSQL ('UPDATE DEVUES SET DV_SQL = SUBSTRING(DV_SQL,1,LEN(CAST(DV_SQL AS VARCHAR(MAX)))-1) WHERE DV_NOMVUE="PGLANCEBULLSANSPAIE" AND ASCII(SUBSTRING(DV_SQL,LEN(CAST(DV_SQL AS VARCHAR(MAX))),1))=0');
    ExecuteSQL ('UPDATE MSAEVOLUTIONSPE2 SET PE2_LIEUTRAV="" WHERE PE2_LIEUTRAV IS NULL');
    ExecuteSQL ('UPDATE DADS2SALARIES SET PD2_BASEIMPO3=0 WHERE PD2_BASEIMPO3 IS NULL');
    ExecuteSQL ('UPDATE ATTESTATIONS SET PAS_DATREPEANTICIP="01/01/1900", PAS_DATDERNJOUR="01/01/1900", PAS_DATENONREPRIS="01/01/1900" WHERE PAS_DATREPEANTICIP IS NULL');
    AglNettoieListes ('PGLANCEBULLAVEC', 'PPU_DATEENTREE');
    AglNettoieListes ('PGLANCEBULLAVEC', 'PPU_DATESORTIE');
    // Ajout BTP Ed6
    ExecuteSQL('UPDATE DECHAMPS SET DH_CONTROLE="LDC" WHERE DH_NOMCHAMP="GP_DEVENIRPIECE"')
  end;
  if VersionBaseDest < '998.ZT' then // edition 12 et 13  paie et compta
  begin
	  ExecuteSQL('UPDATE DETABLES SET DT_NUMVERSION=103 WHERE DT_NOMTABLE="REGLTIJSS" AND DT_NUMVERSION=104');
	  ExecuteSQL('UPDATE DETABLES SET DT_NUMVERSION=140 WHERE DT_NOMTABLE="SALARIES" AND DT_NUMVERSION=141');
  end;
  if VersionBaseDest < '998.ZY' then // edition 16 paie
  begin
    ExecuteSQL ('delete from menu where mn_1 = "41" and mn_2 = "13" and mn_tag like "%419%"');
    ExecuteSQL ('delete from menu where mn_1 = "48" and mn_2 = "3"');
    ExecuteSQL ('update commun set co_libre = " " where co_type="PZN" and co_libre is null');
  end;

  if VersionBaseDest < '998.ZZ7' then
  begin
    ExecuteSQL('update piedport set gpt_basehtdev=gpt_totalhtdev,gpt_baseht=gpt_totalht,gpt_basettc=gpt_totalttc,'+
              'gpt_basettcdev=gpt_totalttcdev where gpt_typeport in ("MT","MTC")')
  end;
	ExecuteSQL('UPDATE CODECPTA SET GCP_VENTEACHAT="" WHERE GCP_VENTEACHAT IS NULL');
  if VersionBaseDest < '998.ZZQ' then
  begin
    ExecuteSql('update article set ga_fonction='+
               '(select ##TOP 1## ars_fonction1 from ressource where ars_article=ga_article) where '+
               'ga_article in (select distinct ars_article from ressource where ars_article <> "" and ars_fonction1 <> "")');
  end;
  if VersionBaseDest < '998.ZZX' then
  begin
    ExecuteSql('DELETE FROM PARAMSOC WHERE SOC_NOM="LSO_BTECHGPIEENV"');
    ExecuteSql('DELETE FROM PARAMSOC WHERE SOC_NOM="SO_BTECHGPIEENV"');
    ExecuteSql('DELETE FROM PARAMSOC WHERE SOC_NOM="SO_BTECHGPIEREC"');
    ExecuteSql('DELETE FROM PARAMSOC WHERE SOC_NOM="LSO_BTECHGPIEREC"');
    ExecuteSql('DELETE FROM PARAMSOC WHERE SOC_NOM="SCO_BTECHGBOX2"');
  end;
  if VersionBaseDest < '998.ZZZH' then
  begin
    ExecuteSQL('UPDATE PARPIECE SET GPP_APPELPRIX="PAS" WHERE GPP_NATUREPIECEG IN ("CF","CFR","BLF","FF") AND GPP_APPELPRIX="DPA"');
  end;
  if VersionBaseDest < '998.ZZZP' then
  begin
    if ExisteSQL('SELECT 1 FROM BTYPELIGBAST WHERE BM6_CODE="FRA0000000003"') then
    begin
      ExecuteSQL('UPDATE BTYPELIGBAST SET BM6_CODE="FRA0000000003",BM6_CATLIGNE="004",BM6_INDICE=3, BM6_TYPEERP="PTC",BM6_ARTICLE="",BM6_LIBELLE="Pénalités " WHERE BM6_CODE="TRA0000000006"');
    end;
    ExecuteSQL('INSERT INTO BTYPELIGBAST (BM6_CATLIGNE,BM6_INDICE,BM6_CODE,BM6_LIBELLE,BM6_ARTICLE,BM6_SENS,BM6_TYPEERP,BM6_AUTHSUP) VALUES ("007",1,"ESC0000000001","Escompte","","-","PTC","-")');
  end;
  if VersionBaseDest < '998.ZZZR' then
  begin
    ExecuteSQL('UPDATE DECOMBOS SET DO_DOMAINE="B" WHERE DO_COMBO IN ("BTSOUSFAMTARART","BTAFFCREERPAR","BTEMPLACEMENT","BTMODELETAT")');
  end;
  if VersionBaseDest < '998.ZZZT' then
  begin
    ExecuteSQL('UPDATE PARAMSOC SET SOC_DESIGN="M;BTPORT##;Prorata;74;387;227;408;X;-;27;389;112;405;MS Sans Serif;-16777208;8;0;MS Sans Serif;-16777208;8;0;-;0;0;X" WHERE SOC_NOM="SO_PRORATA"');
    ExecuteSQL('UPDATE PARAMSOC SET SOC_DESIGN="M;BTPORT##;Révision;302;387;455;408;X;-;248;389;333;405;MS Sans Serif;-16777208;8;0;MS Sans Serif;-16777208;8;0;-;0;0;X" WHERE SOC_NOM="SO_REVISION"');
  end;
  // A faire systématiquement ---
  ExecuteSQL('UPDATE MENU SET MN_ACCESGRP="'+StringOfChar('0',100)+'" WHERE MN_ACCESGRP IS NULL');
  ExecuteSQL('UPDATE MENU SET MN_ACCESGRP="'+StringOfChar('-',100)+'" WHERE MN_VERSIONDEV="X" AND MN_1<>0');
  ExecuteSQL('UPDATE MENU SET MN_VERSIONDEV="-" WHERE MN_VERSIONDEV="X"');
//  ExecuteSQL('DELETE MENU WHERE MN_TAG="-147800"');
end;

procedure TMajStructBTP.ControleTableEchange;
var SQl : string;
    II : Integer;
begin
  if Not ExisteSql('SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE="BASE TABLE" AND TABLE_NAME="ENTETE_ECHANGETPI"') then
  begin
    SQL := 'CREATE TABLE ENTETE_ECHANGETPI('+
           'TPIE_IDUNIQUE int IDENTITY(1,1) NOT NULL,'+
	         'TPIE_IDTPI varchar(30) NOT NULL,'+
           'TPIE_USER varchar(3) NOT NULL,'+
           'TPIE_TYPEECHG varchar(3) NOT NULL,'+
           'TPIE_DATEECHG datetime NOT NULL,'+
           'TPIE_CHANTIER varchar(18) NULL,'+
           'TPIE_DEPOT varchar(3) NULL'+
           ')';
    II := ExecuteSQL(SQl);
    if II <> 1 then
    begin
      SQl := 'CREATE UNIQUE CLUSTERED INDEX ENTECHGTPI1 ON ENTETE_ECHANGETPI('+
	           'TPIE_IDUNIQUE ASC)';
      ExecuteSQL(SQl);
      SQL := 'CREATE NONCLUSTERED INDEX ENTECHGTPI2 ON ENTETE_ECHANGETPI('+
	           'TPIE_USER ASC,'+
	           'TPIE_TYPEECHG ASC)';
      ExecuteSQL(SQl);
    end;
  end;

  if Not ExisteSql('SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE="BASE TABLE" AND TABLE_NAME="LIGNE_ECHANGETPI"') then
  begin
    SQL := 'CREATE TABLE LIGNE_ECHANGETPI('+
        	'TPIL_IDUNIQUE int NOT NULL,'+
	        'TPIL_CODEARTICLE varchar(18) NOT NULL,'+
	        'TPIL_QTE numeric(19, 4) NOT NULL,'+
	        'TPIL_DATE datetime NOT NULL,'+
	        'TPIL_ANALYTIQUE varchar(15) NULL'+
          ')';

    II := ExecuteSQL(SQl);
    if II <> 1 then
    begin
      SQl := 'CREATE CLUSTERED INDEX LIGECHGTPI1 ON LIGNE_ECHANGETPI('+
             'TPIL_IDUNIQUE ASC)';
      ExecuteSQL(SQl);
    end;
  end;

end;

procedure TMajStructBTP.Updateparamsocs;
var TheParamSocsRef,TheParamSocsDest : TOB;
    TheParamSocRef,TheParamSocDest : TOB;
    indice : integer;
    QQ : TQuery;
begin
  TheParamSocsref := TOB.create ('LES PARAMS SOC',nil,-1);
  TheParamSocsDest := TOB.create ('LES PARAMS SOCD',nil,-1);

	XX.Titre.Caption := 'Mise à niveau des Paramètres Sociétés';
  XX.Titre.visible := true;
  XX.Titre.Refresh;
  XX.PB.Visible := true;
  XX.Refresh;
  GetAllParamSocs (TheParamSocsRef,TheParamSocsDest);
  for indice := 0 to TheParamSocsref.detail.count -1 do
  begin
    TheParamSocRef :=TheParamSocsref.detail[Indice];
    TheParamSocDest := TheParamSocsDest.findFirst(['SOC_NOM'],[TheParamSocRef.GetValue('SOC_NOM')],true);
    if TheParamSocDest = nil then
    begin
      TheParamSocDest := TOB.Create('PARAMSOC',TheParamSocsDest,-1);
      QQ := OpenSql ('SELECT * FROM PARAMSOC WHERE SOC_NOM="'+TheParamSocRef.GetValue('SOC_NOM')+'"',true);
      if QQ.eof then
      begin
        // Nouveau paramsoc
        TheParamSOcDest.Dupliquer (TheParamSocRef,false,true);
        TheParamSocDest.SetAllModifie (true);
      end else
      begin
        // Cas particulier du changement de SOC_TREE
        TheParamSocDest.selectDB('',QQ);
        TheParamSocDest.putValue('SOC_DESIGN',TheParamSocRef.GetValue('SOC_DESIGN'));
        TheParamSocDest.putValue('SOC_TREE',TheParamSocRef.GetValue('SOC_TREE'));
        TheParamSocDest.SetAllModifie (true);
      end;
      ferme (QQ);
    end else
    begin
      TheParamSocDest.putValue('SOC_DESIGN',TheParamSocRef.GetValue('SOC_DESIGN'));
      TheParamSocDest.putValue('SOC_TREE',TheParamSocRef.GetValue('SOC_TREE'));
      TheParamSocDest.SetAllModifie (true);
    end;
  end;
  //
  if TheParamSocsref.detail.count <> 0 then
  begin
    TheParamSocsDest.InsertOrUpdateDB (true);
  end;
  //
  TheParamSocsref.free;
  TheParamSocsDest.free;
  //
end;

procedure TMajStructBTP.GetMenusSpecif(NumeroMenu: string; TheMenusSpecif: TOB);
var Sql,Menus : string;
    QQ : TQUery;
    Mn1,Mn2,Mn3,Mn4 : string;
    OT,DT,TT : TOB;
    II : integer;
begin
  TT := TOB.create ('LES MENUS SPE',nil,-1);
  TRY
    menus := NumeroMenu;
    Mn1 := '';
    Mn2 := '';
    Mn3 := '';
    Mn4 := '';
    //
    Mn1 := READTOKENST(Menus);
    if Menus <> '' then
    begin
      Mn2 := READTOKENST(Menus);
    end;
    if Menus <> '' then
    begin
      Mn3 := READTOKENST(Menus);
    end;
    if Menus <> '' then
    begin
      Mn4 := READTOKENST(Menus);
    end;
    Sql := 'SELECT *,0 AS N1,0 AS N2,0 AS N3, 0 AS N4 FROM MENU WHERE MN_1='+Mn1;
    if Mn2 <> '' then
    begin
      Sql := Sql + ' AND MN_2='+Mn2;
    end;
    if Mn3 <> '' then
    begin
      Sql := Sql + ' AND MN_3='+Mn3;
    end;
    if Mn4 <> '' then
    begin
      Sql := Sql + ' AND MN_4='+Mn4;
    end;
    Sql := SQl + ' AND MN_TAG >= 900000 ORDER BY MN_1,MN_2,MN_3,MN_4';
    QQ := OpenSql (Sql,True);
    TT.LoadDetailDB ('MENU','','',QQ,true);
    ferme(QQ);
    // recherche des papas !!!
    if TT.detail.count = 0 then exit;
    II := 0;
    repeat
      OT := TT.detail[II];
      Dt := nil;
      if OT.GetInteger ('MN_4') > 0 then
      begin
        DT := TheMenusSpecif.findFirst(['MN_1','MN_2','MN_3','MN_4'],
                                            [OT.GetValue('MN_1'),
                                             OT.GetValue('MN_2'),
                                             OT.GetValue('MN_3'),
                                             0],true);
      end else if OT.GetInteger ('MN_3') > 0 then
      begin
        DT := TheMenusSpecif.findFirst(['MN_1','MN_2','MN_3','MN_4'],
                                            [OT.GetValue('MN_1'),
                                             OT.GetValue('MN_2'),
                                             0,0],true);
      end else if OT.GetInteger ('MN_2') > 0 then
      begin
        DT := TheMenusSpecif.findFirst(['MN_1','MN_2','MN_3','MN_4'],
                                            [OT.GetValue('MN_1'),
                                             0,0,0],true);
      end;
      if Dt <> nil then
      begin
        OT.ChangeParent (DT,-1);
      end else
      begin
        OT.ChangeParent (TheMenusSpecif,-1);
      end;
    until II >= TT.detail.count;
  FINALLY
  TT.free;
  END;
end;

procedure TMajStructBTP.AddmenusSpecif (TheMenusRef,TheMenusSpecif : TOB);

  procedure SetNewNumFille (OT : TOB);
  var II : integer;
  begin
    for II := 0 to OT.detail.count -1 do
    begin
      if OT.getInteger('N3')> 0 then
      begin
        OT.detail[II].SetInteger ('N4',OT.getInteger('N4')+1);
        OT.detail[II].SetInteger ('N3',OT.getInteger('N3'));
        OT.detail[II].SetInteger ('N2',OT.getInteger('N2'));
        OT.detail[II].SetInteger ('N1',OT.getInteger('N1'));
      end else if OT.getInteger('N2')> 0 then
      begin
        OT.detail[II].SetInteger ('N4',OT.getInteger('N4'));
        OT.detail[II].SetInteger ('N3',OT.getInteger('N3')+1);
        OT.detail[II].SetInteger ('N2',OT.getInteger('N2'));
        OT.detail[II].SetInteger ('N1',OT.getInteger('N1'));
      end else if OT.getInteger('N1')> 0 then
      begin
        OT.detail[II].SetInteger ('N4',OT.getInteger('N4'));
        OT.detail[II].SetInteger ('N3',OT.getInteger('N3'));
        OT.detail[II].SetInteger ('N2',OT.getInteger('N2')+1);
        OT.detail[II].SetInteger ('N1',OT.getInteger('N1'));
      end;
      if OT.detail[II].detail.count > 0 then SetNewNumFille(OT.detail[II]);
    end;
  end;

  procedure TrouveMenuStdLibre (TheMenusRef,OT: TOB; ID : integer);
  var I1,I2,I3,I4 : integer;
      II,Niv : integer;
  begin
    I1 := OT.GetInteger ('MN_1');
    I2 := OT.GetInteger ('MN_2');
    I3 := OT.GetInteger ('MN_3');
    I4 := OT.GetInteger ('MN_4');
    if I4 > 0 then Niv := 3  else if I3 > 0 then Niv := 2 else if I2 > 0 then Niv := 1;
    For II := ID to TheMenusRef.detail.count -1 do
    begin
      if (Niv = 3) and ( (TheMenusRef.detail[II].getInteger('MN_3')<>I3) OR
                         (TheMenusRef.detail[II].getInteger('MN_2')<>I2) OR
                         (TheMenusRef.detail[II].getInteger('MN_1')<>I1) ) then break
      else if (Niv = 2) and ( (TheMenusRef.detail[II].getInteger('MN_2')<>I2) OR
                         (TheMenusRef.detail[II].getInteger('MN_1')<>I1) ) then break
      else if (Niv = 1) and ((TheMenusRef.detail[II].getInteger('MN_1')<>I1)) then break;

      //
      OT.setInteger('N1', TheMenusRef.detail[II].getInteger('MN_1'));
      OT.setInteger('N2', TheMenusRef.detail[II].getInteger('MN_2'));
      OT.setInteger('N3', TheMenusRef.detail[II].getInteger('MN_3'));
      OT.setInteger('N4', TheMenusRef.detail[II].getInteger('MN_4'));
    end;
    if niv = 3 then
    begin
      OT.SetInteger ('N4',OT.getInteger('N4')+1);
    end else if niv = 2 then
    begin
      OT.SetInteger ('N3',OT.getInteger('N3')+1);
    end else if niv = 1 then
    begin
      OT.SetInteger ('N2',OT.getInteger('N2')+1);
    end;
    if OT.detail.count > 0 then SetNewNumFille (OT);
  end;

  procedure InsereEltSpecif (DT,TheMenusRef : TOB);
  var II : integer;
      NT : TOB;
  begin
    NT := TOB.Create ('MENU',TheMenusref,-1);
    NT.Dupliquer(DT,false,true);
    NT.SetInteger('MN_1',NT.GetInteger('N1'));
    NT.SetInteger('MN_2',NT.GetInteger('N2'));
    NT.SetInteger('MN_3',NT.GetInteger('N3'));
    NT.SetInteger('MN_4',NT.GetInteger('N4'));
    if DT.detail.count > 0 then
    begin
      for II := 0 to DT.detail.count -1 do
      begin
        InsereEltSpecif (DT.detail[II],TheMenusRef);
      end;
    end;
  end;

var II,ID : integer;
    OT,DT : TOB;
begin
  for II := 0 to TheMenusSpecif.detail.count -1 do
  begin
    ID := 0;
    OT := TheMenusSpecif.detail[II];
    if OT.GetInteger ('MN_4') > 0 then
    begin
      DT := TheMenusRef.findFirst(['MN_1','MN_2','MN_3','MN_4'],
                                  [OT.GetValue('MN_1'),
                                   OT.GetValue('MN_2'),
                                   OT.GetValue('MN_3'),
                                  0],true);
    end else if OT.GetInteger ('MN_3') > 0 then
    begin
      DT := TheMenusRef.findFirst(['MN_1','MN_2','MN_3','MN_4'],
    															[OT.GetValue('MN_1'),
                                  OT.GetValue('MN_2'),
                                  0,0],true);
    end else if OT.GetInteger ('MN_2') > 0 then
    begin
      DT := TheMenusRef.findFirst(['MN_1','MN_2','MN_3','MN_4'],
    															[OT.GetValue('MN_1'),
                                  0,0,0],true);
    end else continue;
    if DT <> nil then ID := DT.GetIndex ;
    TrouveMenuStdLibre (TheMenusRef,OT,ID);
    InsereEltSpecif (OT,TheMenusRef);
  end;
end;

procedure TMajStructBTP.InitChampsUniqueBAST;
var QQ : TQuery;
    II : Integer;
begin
  II := 1;
  QQ := OpenSQL('SELECT * from BTFACTST',False,-1,'');
  if not QQ.eof then
  begin
    QQ.First;
    repeat
      QQ.Edit;
      QQ.FindField('BM3_UNIQUE').AsInteger := II;
      QQ.Post;
      QQ.Next;
      Inc(II);
    until QQ.Eof;
  end;
  Ferme(QQ);
end;

end.
