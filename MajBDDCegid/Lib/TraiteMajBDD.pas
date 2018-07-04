{***********UNITE*************************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 04/07/2003
Modifié le ... :   /  /
Description .. : - Ajout le 04-07-2003 de la méthode OnAfterFormShow à)
Suite ........ : l'identique de celui du mul.   formule
Mots clefs ... :
*****************************************************************}
unit TraiteMajBDD;
{$IFDEF BASEEXT}dfg{$ENDIF}

interface

uses
  Windows,
  Messages,
  SysUtils,
  Classes,
  Graphics,
  Controls,
  Forms,
  Dialogs,
  {$IFDEF EAGLCLIENT}
  {$ELSE}
  HDimension,
  QRE,
  {$IFDEF ODBCDAC}
  odbcconnection,
  odbctable,
  odbcquery,
  odbcdac,
  {$ELSE}
  {$IFNDEF DBXPRESS}dbtables,
  {$ELSE}uDbxDataSet,
  {$ENDIF}
  {$ENDIF}
  {$ENDIF}
  ComCtrls,
  HSysMenu,
  {$IFNDEF EAGLSERVER}
  AglInit,
  {$ENDIF EAGLSERVER}
  UIUtil,
  HPanel,
  HCtrls,
  Hent1,
  UTOB,
  HTB97,
  IniFiles,
  HMsgBox,
 	MajTable,
 	HDB,
  UTOF, StdCtrls;

type
  //
  ETraitErreur = class (Exception);

  TFMAJBDDBTP = class(THForm)
    HMTrad: THSystemMenu;
    Dock971: TDock97;
    PBouton: TToolWindow97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    bDefaire: TToolbarButton97;
    Binsert: TToolbarButton97;
    BDelete: TToolbarButton97;
    BImprimer: TToolbarButton97;
    BCSRC: TComboBox;
    LCBDEST: TLabel;
    BLANCETRAIT: TToolbarButton97;
    CURVUE: THLabel;
    Mresult: TMemo;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure HelpBtnClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure bDefaireClick(Sender: TObject);
    procedure BinsertClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure BFermeClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BLANCETRAITClick(Sender: TObject);
  private { Déclarations privées }
  	//
    OkContinue : integer;
    ErrorMsg : string;
    DBRef : TDatabase;
    DestDriver: TDBDriver;
    DestODBC: Boolean;
    //
    RefDriver : TDBDriver;
    RefODBC : Boolean;
    //
    IniFile : Tinifile;
    FTOFName: string;
    WMinX, WMinY: Integer;
    FClosing, FFormResize: Boolean;
    FTypeAction: TActionFiche;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    procedure SetFormResize(const Value: Boolean);
    procedure MajBDD;
    procedure TraiteTables;
    procedure TraiteVues;
    procedure Traitelistes;
    procedure TraiteFiches;
    procedure TraiteTablettes;
    procedure TraiteParamsocs;
    procedure TraiteEtats;
    procedure TraiteMenus;
    procedure TraiteGraphs;

    //
    procedure CompleteUneTable (NomTable : string);
    procedure RemplaceUneTable (NomTable : string);

    procedure CreeTable(TheStructureRef: TOB; NomTable: string);
    procedure GetStructureTableFromRef(TheSTructureRef: TOB;NomTable: string);
    procedure GetStructureTable(TheSTructure: TOB; NomTable: string);
    procedure CompleteStructure(TheSTructure, TheSTructureRef: TOB);
    function ModifieTable(TheStructureRef: TOB; NomTable: string): boolean;
    procedure CopieDonnes(NomTable: string);
    procedure TraiteTabletteDetail(TOBTab: TOB);
    //
  public { Déclarations publiques }
    { Nature et code de la fiche }
    FCodeNature, FCodeFiche: string;
    FRetour: hString;
    FArgument: string;
    LaTOF: TOF;
    FAGL: Boolean;
    FMulQ: TQuery;
    OnAfterFormShow: Tproc;
  published
    property TOFName: string read FTOFName write FTOFName;
    property FormResize: Boolean read FFormResize write SetFormResize;
    property Retour: hstring read FRetour write FRetour stored false;
    property Argument: string read FArgument write FArgument stored false;
    property TypeAction: TActionFiche read FTypeAction write FTypeAction;
  end;

const TEMPTABLE = 'TEMPOR';

procedure LanceTraitementActualisationBDD;

implementation

{$R *.DFM}

procedure LanceTraitementActualisationBDD;
var XX : TFMAJBDDBTP;
begin
	XX := TFMAJBDDBTP.Create(Application);
  TRY
  	XX.ShowModal;
  finally
    XX.Free;
  end;
end;

procedure TFMAJBDDBTP.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  IniFile.Free;
  TheFormVM.CloseForm(Self);
  if IsInside(Self) then Action := caFree;
  FClosing := TRUE;
end;

procedure TFMAJBDDBTP.FormShow(Sender: TObject);
var fic : string;
begin
  if V_PGI.AGLDesigning then exit;
  if (not V_PGI.AGLDesigning) then
    LaTOF := CreateTOF(TOFName, Self, TRUE, FAGL)
  else
    LaTof := nil;
  if LaTOF <> nil then LaTOF.LaTOB := TheTOB;
  TheTOB := nil;
  if LaTOF <> nil then LaTOF.Argument(FArgument);
  if FAGL then LanceM3EventProc(Self, 'TOF_OnArgument', [FArgument]); // XP
  if LaTOF <> nil then LaTOF.load;
  if FAGL then LanceM3EventProc(Self, 'TOF_OnLoad', [0]); // XP

//EPZ 15.01.2007
  WMinX := Width;
  WMinY := Height;

  { XP }
  AglEmpileFiche(Self);
  fic := IncludeTrailingBackslash(ExtractFilePath(Application.ExeName))+'MigBDD.ini';
  IniFile := TIniFile.create (fic);
  iniFile.ReadSections(BCSRC.Items);
  if assigned(OnAfterFormShow) then OnAfterFormShow;
end;

procedure TFMAJBDDBTP.BValiderClick(Sender: TObject);
begin
  if not nextPrevControl(Self, false, true) then {PMJ560C}
  begin
    ModalResult := mrNone;
    exit;
  end;
  if LaTOF <> nil then LaTOF.Update;
  if FAGL then LanceM3EventProc(Self, 'TOF_OnUpdate', [0]); // XP
end;

procedure TFMAJBDDBTP.HelpBtnClick(Sender: TObject);
begin
  CallHelpTopic(Self);
end;

procedure TFMAJBDDBTP.FormCreate(Sender: TObject);
begin
  FAGL := FALSE;
  LaTOF := nil;
  FMulQ := nil;
  FClosing := FALSE;
  OnAfterFormShow := nil;
//EPZ 15.01.2007
//EPZ 15.01.2007  WMinX := Width;
//EPZ 15.01.2007  WMinY := Height;
end;

procedure TFMAJBDDBTP.FormDestroy(Sender: TObject);
begin
  if LaTOF <> nil then LaTOF.Free;
  { XP }
  AglDepileFiche;
end;

procedure TFMAJBDDBTP.BDeleteClick(Sender: TObject);
begin
  if LaTOF <> nil then LaTOF.Delete;
  if FAGL then LanceM3EventProc(Self, 'TOF_OnDelete', [0]); // XP
end;

procedure TFMAJBDDBTP.bDefaireClick(Sender: TObject);
begin
  if LaTOF <> nil then LaTOF.Cancel;
  if FAGL then LanceM3EventProc(Self, 'TOF_OnCancel', [0]); // XP
end;

procedure TFMAJBDDBTP.BinsertClick(Sender: TObject);
begin
  if LaTOF <> nil then LaTOF.New;
  if FAGL then LanceM3EventProc(Self, 'TOF_OnNew', [0]); // XP
end;

procedure TFMAJBDDBTP.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  if LaTOF <> nil then CanClose := LaTOF.Close;
  if FAGL and CanClose then CanClose := (LanceM3EventFunc(Self, 'TOF_OnClose', [0]) = 1); // XP
end;

procedure TFMAJBDDBTP.WMGetMinMaxInfo(var MSG: Tmessage);
begin
  inherited;
  exit; //EPZ 14.02.2007 laisser l'utilisateur redimensionner l'écran
  if V_PGI.AGLDesigning then exit; //EPZ 15.01.2007
  with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do
  begin
    X := WMinX;
    Y := WMinY;
  end;
end;

procedure TFMAJBDDBTP.SetFormResize(const Value: Boolean);
begin
  FFormResize := Value;
  HMTrad.ActiveResize := Value;
end;

procedure TFMAJBDDBTP.BFermeClick(Sender: TObject);
begin
  Close;
  if FClosing and IsInside(Self) then THPanel(parent).CloseInside;
end;

procedure TFMAJBDDBTP.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (Key = VK_F10) and (BValider.Visible) and (BValider.Enabled) then
  begin key := 0;
    BValider.Click;
  end;
  if (Key = vk_Imprime) and (ssCtrl in Shift) then
  begin key := 0;
    BImprimer.Click;
  end;
  if (Key = vk_Nouveau) and (ssCtrl in Shift) and (bInsert.Visible) and (bInsert.Enabled) then
  begin key := 0;
    Binsert.Click;
  end;
  if (Key = vk_defaire) and (ssCtrl in Shift) and (bDefaire.visible) and (bDefaire.Enabled) then
  begin key := 0;
    BDefaire.Click;
  end;
  if (Key = vk_delete) and (ssCtrl in Shift) and (bDelete.visible) and (bDelete.Enabled) then
  begin key := 0;
    bDelete.Click;
  end;
end;

procedure TFMAJBDDBTP.BLANCETRAITClick(Sender: TObject);
begin
  if BCSRC.Text = V_PGI.DBName then
  begin
    Application.MessageBox('Sélectionnez une autre base que celle ou vous êtes connecté','Contrôle');
    exit;
  end;
	if Application.MessageBox(Pchar('Etes vous sur de lancer le traitement de mise à jour depuis '+BCSRC.Text+' ?'),'Mise à jour',MB_OKCANCEL	)= IDOK	 then
  begin
     MajBDD;
  end;
end;

procedure TFMAJBDDBTP.MajBDD;
var JJ : integer;
begin
  DestDriver := V_PGI.Driver; DestODBC := V_PGI.ODBC; V_PGI.StopCourrier := TRUE;
  //
  Mresult.Clear;
  JJ := Mresult.Lines.Add('Connection a '+BCSRC.Text);
  DBRef := TDatabase.Create(Application);
  DBRef.Name := BCSRC.Text;
  // Overture de la base Destination
  ConnecteDB(BCSRC.Text, DBRef, 'DBRef');
  Mresult.Lines[JJ] := Mresult.lines[JJ] + ' --> OK';
  RefDriver := V_PGI.Driver; RefODBC := V_PGI.ODBC;
  //
  OkContinue := 0;
  TRY
    // Afin d'eviter les problèmes sur les vues stockés --> on les vires
	  JJ := Mresult.Lines.Add('SUPPRESSION DES VUES PHYSIQUES');
    CURVUE.Visible := True;
    DBDeleteAllView (DBSOC,CURVUE,DestDriver);
	  Mresult.Lines[JJ] := Mresult.Lines[JJ]  + ' --> OK';
    CURVUE.Visible := false;
    //
    Mresult.Lines.Add('------------------------');
    Mresult.Lines.Add('---   DEBUT TABLES  ----');
    Mresult.Lines.Add('------------------------');
    TraiteTables;
    if OkContinue = 0 then
    begin
	  	Mresult.Lines.Add('------------------------');
	  	Mresult.Lines.Add('---    DEBUT VUES   ----');
	  	Mresult.Lines.Add('------------------------');
    	TraiteVues;
    end;
    if OkContinue = 0 then
    begin
	  	Mresult.Lines.Add('------------------------');
	  	Mresult.Lines.Add('---  DEBUT LISTES   ----');
	  	Mresult.Lines.Add('------------------------');
    	Traitelistes;
    end;
    if OkContinue = 0 then
    begin
	  	Mresult.Lines.Add('------------------------');
	  	Mresult.Lines.Add('---  DEBUT FICHES   ----');
	  	Mresult.Lines.Add('------------------------');
    	TraiteFiches;
    end;
    if OkContinue = 0 then
    begin
	  	Mresult.Lines.Add('-------------------------');
	  	Mresult.Lines.Add('---  DEBUT TABLETTES ----');
	  	Mresult.Lines.Add('-------------------------');
    	TraiteTablettes;
    end;
    if OkContinue = 0 then
    begin
	  	Mresult.Lines.Add('-------------------------');
	  	Mresult.Lines.Add('---  DEBUT PARAMSOCS ----');
	  	Mresult.Lines.Add('-------------------------');
    	TraiteParamsocs;
    end;
    if OkContinue = 0 then
    begin
	  	Mresult.Lines.Add('-------------------------');
	  	Mresult.Lines.Add('---    DEBUT ETATS   ----');
	  	Mresult.Lines.Add('-------------------------');
    	TraiteEtats;
    end;
    if OkContinue = 0 then
    begin
	  	Mresult.Lines.Add('-------------------------');
	  	Mresult.Lines.Add('---    DEBUT MENUS   ----');
	  	Mresult.Lines.Add('-------------------------');
    	TraiteMenus;
    end;
    if OkContinue = 0 then
    begin
	  	Mresult.Lines.Add('------------------------------------');
	  	Mresult.Lines.Add('---    DEBUT TABLE DES GRAPHS   ----');
	  	Mresult.Lines.Add('------------------------------------');
    	TraiteGraphs;
    end;
    if OkContinue = 0 then
    begin
	  	Mresult.Lines.Add('--------------------------------');
	  	Mresult.Lines.Add('--- CREATION VUES PHYSIQUES ----');
	  	Mresult.Lines.Add('--------------------------------');
    	CURVUE.Visible := True;
    	DBCreateAllView (DBSOC,CURVUE,DestDriver);
    end;
    if OkContinue = 0 then
    begin
    	CURVUE.Visible := false;
    end;
    if OkContinue = 0 then
    begin
	  	Mresult.Lines.Add('----------------------------------------------');
	  	Mresult.Lines.Add('--- POSITIONNEMENT DE LA BASE EN TYPE BTP ----');
	  	Mresult.Lines.Add('----------------------------------------------');
      TRY
      	ExecuteSQL('UPDATE PARAMSOC SET SOC_DATA="014" WHERE SOC_NOM="SO_DISTRIBUTION"')
      except
      	Mresult.Lines.Add('Erreur Typage de la BDD');
      END;
    end;
  FINALLY
  	Mresult.Lines.Add(' ------ DECONNEXION ----------');
    DBRef.Connected := False;
    DBRef.Free;
  end;
  if OkContinue = 0 then Application.MessageBox('Traitement terminé avec succès','Mise à jour base de donnée')
  else Application.MessageBox(PChar('ATTENTION : Une erreur s''est produite pendant le traitement '+ErrorMsg),'Mise à jour base de donnée')
  //
end;


procedure TFMAJBDDBTP.TraiteTables;
var QQ : TQuery;
    LaTable : string;
begin
  LaTable := '';
  ErrorMsg := '';
  QQ := OpenSQLDb(DBRef ,'SELECT DT_NOMTABLE FROM DETABLES WHERE DT_DOMAINE IN ("B","A","G","H","W")',True,-1,'',True);
	try
    if not QQ.eof then
    begin
      QQ.First;
      repeat
        TRY
          LaTable := QQ.findfield('DT_NOMTABLE').AsString;
        	RemplaceUneTable(laTable);
        except
          raise;
        end;
        QQ.Next;
      until QQ.Eof;
    end;
  Finally
    Ferme(QQ);
  end;
  QQ := OpenSQLDb(DBRef ,'SELECT DT_NOMTABLE FROM DETABLES WHERE DT_DOMAINE="R"',True,-1,'',True);
	try
    if not QQ.eof then
    begin
      QQ.First;
      repeat
        TRY
          LaTable := QQ.findfield('DT_NOMTABLE').AsString;
        	CompleteUneTable(laTable);
        except
          raise;
        end;
        QQ.Next;
      until QQ.Eof;
    end;
  Finally
    Ferme(QQ);
  end;
end;

procedure TFMAJBDDBTP.TraiteEtats;
var QQ : TQuery;
    TOBETAT : TOB;
    II,JJ : Integer;
    cledata,Etat : string;
begin
  ExecuteSql('DELETE FROM MODELES WHERE MO_NATURE IN ("GPJ","GFA","GZF","GCE","GZL","INT","HOT","ACE","MST","ART","GEA","GED","GZH","RTA","NAT","RPF","BP1")');
  TOBETAT := TOB.Create('LES FORMES',nil,-1);
  ErrorMsg := '';
  QQ := OpenSQLDb(DBRef ,'SELECT * FROM MODELES WHERE MO_NATURE IN ("GPJ","GFA","GZF","GCE","GZL","INT","HOT","ACE","MST","ART","GEA","GED","GZH","RTA","NAT","RPF","BP1")',True,-1,'',True);
	try
    if not QQ.eof then
    begin
      TOBETAT.loadDetailDb ('MODELES','','',QQ,false);
    end;
  Finally
    Ferme(QQ);
  end;
  if TOBETAT.Detail.count > 0 then
  begin
    for II := 0 to TOBETAT.detail.count -1 do
    begin
      TOBETAT.detail[II].putvalue('MN_DOMAINE','B');
      cledata := TOBETAT.detail[II].GetString('MO_TYPE')+
      					 TOBETAT.detail[II].GetString('MO_NATURE')+
      					 TOBETAT.detail[II].GetString('MO_CODE')+'%'+
      					 TOBETAT.detail[II].GetString('MO_LANGUE');
      Etat := TOBETAT.detail[II].GetString('MO_TYPE')+
      					 TOBETAT.detail[II].GetString('MO_NATURE')+
      					 TOBETAT.detail[II].GetString('MO_CODE');

  		ExecuteSQL('DELETE FROM MODEDATA WHERE MD_CLE LIKE "'+Cledata+'"');
  		QQ := OpenSQLDb(DBRef ,'SELECT * FROM MODEDATA WHERE MD_CLE LIKE "'+Cledata+'"',True,-1,'',True);
      if not QQ.eof then
      begin
				TRY
          TOBETAT.detail[II].LoadDetailDB('MODEDATA','','',QQ,false);
          JJ := Mresult.Lines.add ('Traitement de l''etat '+Etat);
          TOBETAT.detail[II].SetAllModifie (True);
          TOBETAT.detail[II].InsertDB(nil);
          Mresult.Lines[JJ] := Mresult.Lines[JJ]  + ' --> OK ';
        EXCEPT
          on E: Exception do
          begin
            Mresult.Lines[JJ] := Mresult.Lines[JJ]  + ' < ERREUR > ';
            raise;
          end;
        END;
      end;
      ferme (QQ);
    end;
  end;
  TOBETAT.Free;

end;

procedure TFMAJBDDBTP.TraiteFiches;
var QQ : TQuery;
    TOBDFM,TOBSCRIPTS : TOB;
    II,JJ : Integer;
begin
  ExecuteSql('DELETE FROM FORMES WHERE DFM_TYPEFORME IN ("AFF","BTP","GC","H","MBO","RT","W")'); // les dfm
  ExecuteSql('DELETE FROM SCRIPTS WHERE SRC_TYPESCRIPT IN ("AFF","BTP","GC","H","MBO","RT","W")'); // les scripts
  TOBDFM := TOB.Create('LES FORMES',nil,-1);
  TOBSCRIPTS := TOB.Create ('LES SCRIPTS',nil,-1);
  ErrorMsg := '';
  QQ := OpenSQLDb(DBRef ,'SELECT * FROM FORMES WHERE DFM_TYPEFORME IN ("AFF","BTP","GC","H","MBO","RT","W")',True,-1,'',True);
	try
    if not QQ.eof then
    begin
      TOBDFM.loadDetailDb ('FORMES','','',QQ,false);
    end;
  Finally
    Ferme(QQ);
  end;

  QQ := OpenSQLDb(DBRef ,'SELECT * FROM SCRIPTS WHERE SRC_TYPESCRIPT IN ("AFF","BTP","GC","H","MBO","RT","W")',True,-1,'',True);
	try
    if not QQ.eof then
    begin
      TOBSCRIPTS.loadDetailDb ('SCRIPTS','','',QQ,false);
    end;
  Finally
    Ferme(QQ);
  end;

  if TOBDFM.Detail.count > 0 then
  begin
    for II := 0 to TOBDFM.detail.count -1 do
    begin
			TRY
        TOBDFM.detail[II].putvalue('DFM_DOMAINE','B');
        JJ := Mresult.Lines.Add('Traitement de la fiche '+TOBDFM.detail[II].getValue('DFM_FORME'));
        TOBDFM.detail[II].SetAllModifie (True);
        TOBDFM.detail[II].InsertDB(nil);
        Mresult.Lines[JJ] := Mresult.Lines[JJ]  + ' --> OK';
      EXCEPT
        on E: Exception do
        begin
        	Mresult.Lines[JJ] := Mresult.Lines[JJ]  + ' < ERREUR > ';
          raise;
        end;
      END;
    end;
  end;

  if TOBSCRIPTS.Detail.count > 0 then
  begin
    for II := 0 to TOBSCRIPTS.detail.count -1 do
    begin
			TRY
        JJ := Mresult.Lines.Add('Traitement du script de la fiche '+TOBSCRIPTS.detail[II].getValue('SRC_SCRIPT'));
		    TOBSCRIPTS.detail[II].SetAllModifie (True);
        TOBSCRIPTS.detail[II].SetAllModifie (True);
        TOBSCRIPTS.detail[II].InsertDB(nil);
        Mresult.Lines [JJ] := Mresult.Lines [JJ] + ' --> OK';
      EXCEPT
        on E: Exception do
        begin
        	Mresult.Lines [JJ] := Mresult.Lines [JJ] + ' < ERREUR >';
          raise;
        end;
      END;
    end;
  end;

  TOBDFM.Free;
  TOBSCRIPTS.Free;
end;

procedure TFMAJBDDBTP.Traitelistes;
var QQ : TQuery;
    TOBLISTES : TOB;
    II,JJ : integer;
begin
  ExecuteSql('DELETE FROM LISTE WHERE LI_LISTE LIKE ("A%") OR LI_LISTE LIKE ("B%") OR LI_LISTE LIKE ("G%") OR LI_LISTE LIKE ("H%") OR LI_LISTE LIKE ("RT%")');
  TOBLISTES := TOB.Create('LES LISTES',nil,-1);
  ErrorMsg := '';
  QQ := OpenSQLDb(DBRef ,'SELECT * FROM LISTE WHERE LI_LISTE LIKE ("A%") OR LI_LISTE LIKE ("B%") OR LI_LISTE LIKE ("G%") OR LI_LISTE LIKE ("H%") OR LI_LISTE LIKE ("RT%")',True,-1,'',True);
	try
    if not QQ.eof then
    begin
      TOBLISTES.loadDetailDb ('LISTE','','',QQ,false);
    end;
  Finally
    Ferme(QQ);
  for II := 0 to TOBLISTES.detail.count -1 do
  begin
    TRY
      JJ := Mresult.Lines.Add('Traitement de la liste '+TOBLISTES.detail[II].GetString('LI_LISTE'));
      TOBLISTES.detail[II].SetAllModifie (True);
      TOBLISTES.detail[II].InsertDB(nil);
      Mresult.Lines[JJ] := Mresult.Lines[JJ]  + ' --> OK';
    except
      on E: Exception do
      begin
      	Mresult.Lines[JJ] := Mresult.Lines[JJ]  + ' < ERREUR >';
        raise;
      end;
    end;
  end;
//    TOBLISTES.SetAllModifie (True);
//    TOBLISTES.InsertDB(nil);
    TOBLISTES.Free;
  end;
end;

procedure TFMAJBDDBTP.TraiteMenus;
var TOBMenus : TOB;
		II : Integer;
    QQ : TQuery;
begin
// nettoyage
  ExecuteSql('DELETE FROM MENU WHERE MN_1=0 and MN_2 IN (60,148,323,304,284,92,146,147,150,146,283,320,149,319,321,322,145,325,327)');
  ExecuteSql('DELETE FROM MENU WHERE MN_1 IN (60,148,323,304,284,92,146,147,150,146,283,320,149,319,321,322,145,325,327)');
//
  TOBMenus := TOB.Create('LES MENU',nil,-1);
  ErrorMsg := '';
  QQ := OpenSQLDb(DBRef ,'SELECT * FROM MENU WHERE MN_1=0 and MN_2 IN (60,148,323,304,284,92,146,147,150,146,283,320,149,319,321,322,145,325,327)',True,-1,'',True);
	try
    if not QQ.eof then
    begin
      TOBMenus.loadDetailDb ('MENU','','',QQ,false);
    end;
  Finally
    Ferme(QQ);
  end;
  QQ := OpenSQLDb(DBRef ,'SELECT * FROM MENU WHERE MN_1 IN (60,148,323,304,284,92,146,147,150,146,283,320,149,319,321,322,145,327)',True,-1,'',True);
	try
    if not QQ.eof then
    begin
      TOBMenus.loadDetailDb ('MENU','','',QQ,true);
    end;
  Finally
    Ferme(QQ);
  end;
  if TOBMenus.Detail.count > 0 then
  begin
    for II := 0 to TOBMenus.detail.count -1 do
    begin
      TOBMenus.detail[II].putvalue('MN_VERSIONDEV','-');
    end;
    TOBMenus.SetAllModifie (True);
    TOBMenus.InsertDB(nil);
  end;
  TOBMenus.Free;
end;

procedure TFMAJBDDBTP.TraiteParamsocs;
var QQ : TQuery;
		TOBParamS : TOB;
    II,JJ : integer;
begin
  TOBParamS := TOB.Create('LES PARAMSOCS',nil,-1);
  //
  ExecuteSql('DELETE FROM PARAMSOC WHERE SOC_TREE LIKE "001;003;%"');
  ExecuteSql('DELETE FROM PARAMSOC WHERE SOC_TREE LIKE "001;004;%"');
  ExecuteSql('DELETE FROM PARAMSOC WHERE SOC_TREE LIKE "001;007;%"');
  ExecuteSql('DELETE FROM PARAMSOC WHERE SOC_TREE LIKE "001;008;%"');
  ExecuteSql('DELETE FROM PARAMSOC WHERE SOC_TREE LIKE "001;016;%"');
  ExecuteSql('DELETE FROM PARAMSOC WHERE SOC_TREE LIKE "001;027;%"');
  TRY
    //
    QQ := OpenSQLDb(DBRef ,'SELECT * FROM PARAMSOC WHERE SOC_TREE LIKE "001;003;%"',True,-1,'',True);
    try
      if not QQ.eof then
      begin
        TOBParamS.loadDetailDb ('PARAMSOC','','',QQ,false);
      end;
    Finally
      Ferme(QQ);
    end;
    QQ := OpenSQLDb(DBRef ,'SELECT * FROM PARAMSOC WHERE SOC_TREE LIKE "001;004;%"',True,-1,'',True);
    try
      if not QQ.eof then
      begin
        TOBParamS.loadDetailDb ('PARAMSOC','','',QQ,true);
      end;
    Finally
      Ferme(QQ);
    end;
    QQ := OpenSQLDb(DBRef ,'SELECT * FROM PARAMSOC WHERE SOC_TREE LIKE "001;007;%"',True,-1,'',True);
    try
      if not QQ.eof then
      begin
        TOBParamS.loadDetailDb ('PARAMSOC','','',QQ,true);
      end;
    Finally
      Ferme(QQ);
    end;
    QQ := OpenSQLDb(DBRef ,'SELECT * FROM PARAMSOC WHERE SOC_TREE LIKE "001;008;%"',True,-1,'',True);
    try
      if not QQ.eof then
      begin
        TOBParamS.loadDetailDb ('PARAMSOC','','',QQ,true);
      end;
    Finally
      Ferme(QQ);
    end;
    QQ := OpenSQLDb(DBRef ,'SELECT * FROM PARAMSOC WHERE SOC_TREE LIKE "001;016;%"',True,-1,'',True);
    try
      if not QQ.eof then
      begin
        TOBParamS.loadDetailDb ('PARAMSOC','','',QQ,true);
      end;
    Finally
      Ferme(QQ);
    end;
    QQ := OpenSQLDb(DBRef ,'SELECT * FROM PARAMSOC WHERE SOC_TREE LIKE "001;027;%"',True,-1,'',True);
    try
      if not QQ.eof then
      begin
        TOBParamS.loadDetailDb ('PARAMSOC','','',QQ,true);
      end;
    Finally
      Ferme(QQ);
    end;
    For II := 0 to TOBPARAMs.detail.count -1 do
    begin
      TOBParamS.detail[II].SetAllModifie (True);
      TRY
        JJ := Mresult.Lines.add ('ECRITURE PARAMSOC : '+ TOBParamS.detail[II].GetString('SOC_NOM'));
        TOBParamS.detail[II].InsertDB(nil);
        Mresult.Lines [JJ] := Mresult.Lines [JJ] + ' --> OK';
      EXCEPT
        on E: Exception do
        begin
          Mresult.Lines [JJ] := Mresult.Lines [JJ] + ' < ERREUR >';
          raise;
        end;
      END;
    end;
  FINALLY
  	TOBParamS.Free;
  END;
end;

procedure TFMAJBDDBTP.TraiteTabletteDetail (TOBTab: TOB);
var TOBCC,TOBCO,TOBYX : TOB;
		TypeTab,Prefixe : string;
    QQ : TQuery;
    JJ : integer;
begin
  //
  JJ := Mresult.Lines.Add('Données de la Tablette '+TOBTab.GetString('DO_COMBO')+' ---> OK');

  TRY
    Prefixe := TOBTAB.GetString('DO_PREFIXE');
    if Pos(prefixe,'CO;CC;YX')> 0 then
    begin
      TypeTab := TOBTab.GetString('DO_TYPE');
      IF Prefixe = 'CO' then
      begin
        TOBCO := TOB.Create ('LES COMMUNS',nil,-1);
        ExecuteSql('DELETE FROM COMMUN WHERE CO_TYPE="'+TypeTab+'"'); // Suppresion dans COMMUN
        QQ := OpenSQLDb(DBRef ,'SELECT * FROM COMMUN WHERE CO_TYPE="'+TypeTab+'"',True,-1,'',True);
        TOBCO.loadDetailDb('COMMUN','','',QQ,False);
        Ferme(QQ);
        if TOBCO.Detail.Count > 0 then
        begin
          TOBCO.SetAllModifie(true);
          TOBCO.InsertDB(nil);
        end;
        TOBCO.free;
      end else if prefixe = 'CC' then
      begin
        TOBCC := TOB.Create ('LES CHOIXCOD',nil,-1);
        ExecuteSql('DELETE FROM CHOIXCOD WHERE CC_TYPE="'+TypeTab+'"'); // Suppresion dans CHOIXCOD
        QQ := OpenSQLDb(DBRef ,'SELECT * FROM CHOIXCOD WHERE CC_TYPE="'+TypeTab+'"',True,-1,'',True);
        TOBCC.loadDetailDb('CHOIXCOD','','',QQ,False);
        Ferme(QQ);
        if TOBCC.Detail.Count > 0 then
        begin
          TOBCC.SetAllModifie(true);
          TOBCC.InsertDB(nil);
        end;
        TOBCC.free;
      end else if prefixe = 'YX' then
      begin
        TOBYX := TOB.Create ('LES CHOIXEXT',nil,-1);
        ExecuteSql('DELETE FROM CHOIXEXT WHERE YX_TYPE="'+TypeTab+'"'); // Suppresion dans CHOIXEXT
        QQ := OpenSQLDb(DBRef ,'SELECT * FROM CHOIXEXT WHERE YX_TYPE="'+TypeTab+'"',True,-1,'',True);
        TOBYX.loadDetailDb('CHOIXEXT','','',QQ,False);
        Ferme(QQ);
        if TOBYX.Detail.Count > 0 then
        begin
          TOBYX.SetAllModifie(true);
          TOBYX.InsertDB(nil);
        end;
        TOBYX.free;
      end;
    end;
    Mresult.Lines [JJ] := Mresult.Lines [JJ] + ' --> OK';
  EXCEPT
    on E: Exception do
    begin
      Mresult.Lines [JJ] := Mresult.Lines [JJ] + ' < ERREUR >';
      raise;
    end;
  END;
end;

procedure TFMAJBDDBTP.TraiteTablettes;
var QQ : TQuery;
    TOBDonnees,TOBTablettes : TOB;
    II,JJ : Integer;
begin
  // particularité
//  ExecuteSql('DELETE FROM COMMUN WHERE CO_TYPE="NTD"'); // définition des type d'etats
  TOBDonnees := TOB.Create('LES DONNEES',nil,-1);
  ErrorMsg := '';
  QQ := OpenSQLDb(DBRef ,'SELECT * FROM COMMUN WHERE CO_TYPE="NTD"',True,-1,'',True);
	try
    if not QQ.eof then
    begin
      TOBDonnees.loadDetailDb ('COMMUN','','',QQ,false);
    end;
  Finally
    Ferme(QQ);
  end;
  TOBDonnees.SetAllModifie (True);
  TOBDonnees.InsertOrUpdateDB;
  TOBDonnees.ClearDetail;
  TOBDonnees.Free;
  //
  TOBTablettes := TOB.Create('LES TABLETTES',nil,-1);
  ExecuteSql('DELETE FROM DECOMBOS WHERE DO_DOMAINE IN ("B","A","G","R","H","W")');
  QQ := OpenSQLDb(DBRef ,'SELECT * FROM DECOMBOS WHERE DO_DOMAINE IN ("B","A","G","R","H","W")',True,-1,'',True);
  if not QQ.eof then
  begin
    TOBTablettes.LoadDetailDB('DECOMBOS','','',QQ,False);
  end;
  ferme (QQ);
  TRY
  	TOBTablettes.detail.sort ('DO_COMBO');
    for II := 0 to TOBTablettes.detail.count -1 do
    begin
      Try
        JJ := Mresult.Lines.Add('Définition Tablette '+TOBTablettes.detail[II].GetString('DO_COMBO')+' ---> OK');
    		TOBTablettes.detail[II].DeleteDB; // précaution contre ces gnafrons de CEGID
    		TOBTablettes.detail[II].InsertDB(nil); // mise en place des définitions
        Mresult.Lines [JJ] := Mresult.Lines [JJ] + ' --> OK';
      except
        on E: Exception do
        begin
        	Mresult.Lines [JJ] := Mresult.Lines [JJ] + ' < ERREUR >';
          raise;
        end;
      end;

      TRY
        JJ := Mresult.Lines.Add('Données ');
      	TraiteTabletteDetail (TOBTablettes.detail[II]);
        Mresult.Lines [JJ] := Mresult.Lines [JJ] + ' --> OK';
      EXCEPT
        on E: Exception do
        begin
        	Mresult.Lines [JJ] := Mresult.Lines [JJ] + ' < ERREUR >';
          raise;
        end;
      END;
    end;

  FINALLY
  	TOBTablettes.free;
  end;
end;

procedure TFMAJBDDBTP.TraiteVues;
var QQ : TQuery;
    TOBVUES : TOB;
    II,JJ : integer;
begin
  ExecuteSql('DELETE FROM DEVUES WHERE DV_DOMAINE IN ("B","A","G","R","H","W")');
  TOBVues := TOB.Create('LES VUES',nil,-1);
  ErrorMsg := '';
  QQ := OpenSQLDb(DBRef ,'SELECT * FROM DEVUES WHERE DV_DOMAINE IN ("B","A","G","R","H","W")',True,-1,'',True);
	try
    if not QQ.eof then
    begin
      TOBVUES.loadDetailDb ('DEVUES','','',QQ,false);
    end;
  Finally
    Ferme(QQ);
  end;
  for II := 0 to TOBVUES.detail.count -1 do
  begin
    TRY
      JJ := Mresult.Lines.Add('Traitement de la vue '+TOBVUES.detail[II].GetString('DV_NOMVUE'));
      TOBVUES.detail[II].SetAllModifie (True);
      TOBVUES.detail[II].InsertDB(nil);
      Mresult.Lines[JJ] := Mresult.Lines[JJ] + ' --> OK';
    except
      on E: Exception do
      begin
      	Mresult.Lines[JJ] := Mresult.Lines[JJ] + ' < ERREUR >';
        raise;
      end;
    end;
  end;
  TOBVUES.Free;
end;

procedure TFMAJBDDBTP.GetStructureTableFromRef(TheSTructureRef: TOB;NomTable: string);
var QQ : TQuery;
begin
  QQ := OpenSQLDb (DBREF,'SELECT * FROM DETABLES WHERE DT_NOMTABLE="'+NomTable+'"',true);
  if QQ.eof then BEGIN OkContinue := 2; Exit; END;
  TheSTructureRef.SelectDB ('',QQ);
  ferme (QQ);
  QQ := OpenSQLDb (DBREF,'SELECT * FROM DECHAMPS WHERE DH_PREFIXE="'+TheSTructureRef.getValue('DT_PREFIXE')+'"',true);
  TheSTructureRef.LoadDetailDB ('DECHAMPS','','',QQ,false);
  ferme (QQ);
end;

procedure TFMAJBDDBTP.GetStructureTable(TheSTructure: TOB;NomTable: string);
var QQ : TQuery;
begin
  QQ := OpenSQLDb (DBSOC,'SELECT * FROM DETABLES WHERE DT_NOMTABLE="'+NomTable+'"',true);
  if QQ.eof then BEGIN OkContinue := 2; Exit; END;
  TheSTructure.SelectDB ('',QQ);
  ferme (QQ);
  QQ := OpenSQLDb (DBSOC,'SELECT * FROM DECHAMPS WHERE DH_PREFIXE="'+TheSTructure.getValue('DT_PREFIXE')+'"',true);
  TheSTructure.LoadDetailDB ('DECHAMPS','','',QQ,false);
  ferme (QQ);
end;

procedure TFMAJBDDBTP.CreeTable (TheStructureRef : TOB;NomTable : string);
begin
  if TableExiste (TEMPTABLE) then DBDeleteTable(DBSOC,DestDriver, TEMPTABLE, false);
  if TableExiste (NomTable) then
  begin
    DBDeleteTable(DBSOC ,DestDriver,NomTable,true); // physique
  end;
  ExecuteSQL('DELETE FROM DETABLES WHERE DT_NOMTABLE="'+NomTable+'"');
  ExecuteSQL('DELETE FROM DECHAMPS WHERE DH_PREFIXE="'+TheStructureRef.GetValue('DT_PREFIXE')+'"');
  DbCreateTable(DBSOC, TheStructureRef, DestDriver, false);
  // Mise à jour de DETABLES ET DECHAMPS
  TheStructureRef.SetAllModifie (true);
  TheSTructureRef.InsertDB (nil);
end;

procedure TFMAJBDDBTP.CopieDonnes (NomTable : string);
begin
	DBCOPYTABLEBM(DBREF, DBSOC, NomTable, NomTable);
end;


procedure TFMAJBDDBTP.RemplaceUneTable(NomTable: string);
var  TheSTructureRef : TOB;
		 JJ : integer;
begin
  if TableExiste (TEMPTABLE) then DBDeleteTable(DBSOC,DestDriver, TEMPTABLE, TRUE);
  JJ := Mresult.Lines.Add('Table '+Nomtable);
  TheSTructureRef := TOB.Create ('DETABLES',nil,-1);
  TRY
    TRY
    	GetStructureTableFromRef (TheSTructureRef,NomTable);
    EXCEPT
      on E: Exception do
      begin
        Mresult.Lines[JJ] := Mresult.Lines[JJ] + 'ERREUR STRUCTURE';
      	raise ;
        exit;
      end;
    END;

    TRY
      // la table existe dans la base de donnée de destination --> on la drop
      CreeTable (TheStructureRef,NomTable);
      CopieDonnes (NomTable);
      Mresult.Lines[JJ] := Mresult.Lines[JJ] + '--> OK';
    except
      on E: Exception do
      begin
      	Mresult.Lines[JJ] := Mresult.Lines[JJ] + '< - ERREUR - >';
      	raise ;
        exit;
      end;
    end;
  FINALLY
  	TheSTructureRef.Free;
  end;
end;

procedure TFMAJBDDBTP.CompleteStructure (TheSTructure,TheSTructureRef : TOB);
var II : Integer;
		LaTOB,NewTOB : TOB;
    lastNum : Integer;
begin
  for II := 0 to TheSTructureRef.detail.count -1 do
  begin
  	LaTOB := TheSTructureRef.detail[II];
    if TheSTructure.FindFirst(['DH_NOMCHAMP'],[LaTOB.GetValue('DH_NOMCHAMP')],true) = nil then
    begin
  		Mresult.Lines.Add('--> Ajout du champ '+LaTOB.GetValue('DH_NOMCHAMP'));
      lastNum := TheSTructure.detail[TheSTructure.detail.count-1].GetValue('DH_NUMCHAMP');
      NewTOB := TOB.Create('DECHAMPS',TheSTructure,-1);
      NewTOB.Dupliquer(LaTOB,False,true);
      NewTOB.PutValue('DH_NUMCHAMP',lastNum+1);
      NewTOB.PutValue('DH_DATECREATION',Date);
      NewTOB.PutValue('DH_MODIF',Date);
    end;
  end;
end;

function TFMAJBDDBTP.ModifieTable (TheStructureRef : TOB;NomTable : string) : boolean;
begin
  result := true;
  if TableExiste (TEMPTABLE) then DBDeleteTable(DBSOC,DestDriver, TEMPTABLE, TRUE);
  // Phase 2 : Création de la table temporaire avec la nouvelle structure
  DbCreateTable(DBSOC, TheStructureRef, DestDriver, TRUE);

  // Phase 3 : Copie de l'ancienne table dans la nouvelle table
  DBCOPYTABLEBM(DBSOC, DBSOC, NomTable, TEMPTABLE);

  // Suppression de DETABLES et DECHAMPS
  ExecuteSql('DELETE FROM DECHAMPS WHERE DH_PREFIXE="' +TheStructureRef.GetValue('DT_PREFIXE') + '"');
  ExecuteSql('DELETE FROM DETABLES WHERE DT_NOMTABLE="' +TheStructureRef.GetValue('DT_NOMTABLE') + '"');

  // Phase 4 : Suppression de l'ancienne table
  DBDeleteTable(DBSOC, DestDriver, NomTable, TRUE);

  // Phase 5 : Renommage de la table temporaire sous le nouveau nom
  DBRenameTable(DBSOC, DestDriver, TEMPTABLE, NomTable);

  // Phase 6 : Création des index
  //  ConnecteTRUEFALSE(DBSOC);
  DBCREATEINDEX(DBSOC, TheStructureRef, DestDriver);
  // Mise à jour de DETABLES ET DECHAMPS
  TheStructureRef.SetAllModifie (true);
  TheSTructureRef.InsertDB (nil);
end;

procedure TFMAJBDDBTP.CompleteUneTable(NomTable: string);
var  TheSTructureRef,TheStructure : TOB;
begin
  Mresult.Lines.Add('Completion de la table '+Nomtable);
  TheSTructureRef := TOB.Create ('DETABLES',nil,-1);
  TheSTructure := TOB.Create ('DETABLES',nil,-1);
  TRY
    GetStructureTableFromRef (TheSTructureRef,NomTable);
    if OkContinue > 0 then
    begin
      raise ETraitErreur.Create('Récupération de la structure de la table (ref)');
    	Exit;
    end;
    GetStructureTable (TheSTructure,NomTable);
    if OkContinue > 0 then
    begin
      raise ETraitErreur.Create('Récupération de la structure de la table (dest)');
    	Exit;
    end;
    if TheStructureref.getValue('DT_NUMVERSION')>TheStructure.getValue('DT_NUMVERSION') then
    begin
      TheStructure.PutValue('DT_NUMVERSION',TheStructureref.getValue('DT_NUMVERSION'));
    end;
    CompleteStructure (TheSTructure,TheSTructureRef);
    TRY
      ModifieTable (TheStructure,NomTable);
  		Mresult.Lines.Add('--> Table '+nomtable+' OK');
    except
      OkContinue := 3;
      raise ETraitErreur.Create('Modification de la table');
    end;
  FINALLY
  	TheSTructureRef.Free;
  end;
end;

procedure TFMAJBDDBTP.TraiteGraphs;
var TOBgraphs : TOB;
		II,JJ : Integer;
    QQ : TQuery;
begin
// nettoyage
  ExecuteSql('DELETE FROM GRAPHS WHERE GR_GRAPHE LIKE "TBT%"');
//
  TOBgraphs := TOB.Create('LES GRAPHS',nil,-1);
  TRY
    ErrorMsg := '';
    QQ := OpenSQLDb(DBRef ,'SELECT * FROM GRAPHS WHERE GR_GRAPHE LIKE "TBT%"',True,-1,'',True);
    try
      if not QQ.eof then
      begin
        TOBgraphs.loadDetailDb ('MENU','','',QQ,false);
      end;
    Finally
      Ferme(QQ);
    end;
    if TOBgraphs.Detail.count > 0 then
    begin
      for II := 0 to TOBgraphs.detail.count -1 do
      begin
        JJ := Mresult.Lines.Add('Traitement du graphe '+TOBgraphs.detail[II].Getvalue('GR_GRAPHE'));
        TOBgraphs.detail[II].SetAllModifie (True);
      TRY
        TOBgraphs.detail[II].InsertDB(nil);
        Mresult.lines[JJ] := Mresult.lines[JJ]  + ' --> OK';
      except
        Mresult.lines[JJ] := Mresult.lines[JJ]  + ' < ERREUR >';
        OkContinue := 3;
        raise;
        exit;
      end;
      end;
    end;
  FINALLY
  	TOBgraphs.Free;
  END;
end;

initialization
  registerclasses([TFMAJBDDBTP]);
end.

