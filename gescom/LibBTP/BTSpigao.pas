unit BTSpigao;

interface

uses
  Windows, Classes, SysUtils, HMsgBox, HCtrls, FE_Main,   controls,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  uTOB, ParamSoc, HEnt1, UtilArticle, GerePiece, FactComm, CBPPath, Web,
  Edisys_IULM_Core_TLB ,Edisys_IULM_CoreEvents,Forms;

type

  T_SPIGAOCONNECT = class (TObject)
  private
    TOBPiece,TOBLignes,TOBDETETUDE : TOB;
    ArticleDivers : string;
    fLoader : ISpigao;
    fDetect : IDetectionEngine;
    fXX : TEdisys_IULM_CoreIDetectionEvents2;
    fconnected : boolean;
    fVersion : string;
    fOldGestion : Boolean;
    fNewgestion : Boolean;
    function ChargeProjet (Direct : Boolean; CodeAffaire : String; IdSpigao : double = 0 ) : IProject;
    Procedure ChargeTob(currentCollection : IWorkItemCollection; Niv : Integer=0);
    procedure UserRequestNewDeals (Sender : TObject);
    procedure UserRequestUpdatedDeals (Sender : TObject);
    procedure UserRequestDealsWithAvailableDIE (Sender : TObject);
  public
    property connected : boolean read fconnected;
    property IsNewGestion : boolean read fNewgestion;
    constructor create;
    destructor destroy; override;
    // Connection - deconnexion
    function Connecte : Boolean;
    procedure deconnecte;
    // --
    function ChargeAO (CodeAffaire, CodeTiers, CodeEtab : String;IdSPIGAO : double) : Boolean;
    function ChargeAffaire (Var MonAffaire : IDeal; FromNew : boolean) : Boolean;
    Procedure MajStatusAffaire (MonAffaireId : Integer);
    function ExportAO (CodeAffaire : String) : Boolean;
    function VoirAffaire (Id : Integer) : Boolean;
    procedure SetGlobalSettings;
    function TelechargerDIE (Id : Integer) : Boolean;
    function TestAffaire (CodeAffaire : String) : Boolean;
    function VoirBPX (CodeAffaire, Identifiant : string) : Boolean;
    procedure RecupNbAffaires (Var Nbnew, NbMod, NbImp, NbMes : Integer);
    procedure VoirAffairesNew;
    procedure VoirAffairesMod;
    procedure LienBTP;
    procedure VoirMessages;
    procedure Confdetection;
    procedure ConfTraitement;
    procedure APropos;
    procedure Infos;
    procedure MonteToaster;
    procedure ControleAcces;
    procedure SetEvents;
    procedure ChargeNewAffaires ;
    procedure IntegreDie ;
  end;

var
  SPIGAOOBJ : T_SPIGAOCONNECT;

implementation
uses affaireutil,UTOF_VideInside
{$IFNDEF PGIMAJVER}
{$IFNDEF MOBILITE}
{$IFNDEF POUCHAIN}
{$IFNDEF SPINNAKER}
,MdispBTP
{$ENDIF}
{$ENDIF}
{$ENDIF}
{$ENDIF}
;

{ T_SPIGAOCONNECT }

procedure T_SPIGAOCONNECT.APropos;
begin
  if not fconnected  then Exit;
  fLoader.ShowAbout;
end;

function T_SPIGAOCONNECT.ChargeAffaire(var MonAffaire: IDeal; FromNew : boolean): Boolean;
var dc : IDealCollection;
		ii : IDealResult;
begin
  Result := False;
  if not fconnected  then Exit;
  if fNewgestion then
  begin
  	ControleAcces;
    if FromNew then
    begin
    	ii := fDetect.UI.Deals.ListNewDeals();
    end else
    begin
    	ii := fDetect.UI.Deals.ListSelectedDeals();
    end;
    if ii.Result = ExecResult_Completed then
    begin
      MonAffaire := ii.SelectedDeal;
      Result := true;
    end else
    begin
      PGIBox('Pas d''affaire à importer');
    end;
  end else
  begin
  	dc := fDetect.UI.ShowDealImporter(False);
    if dc.FullCount > 0 then
    begin
      MonAffaire := dc.Item[0];
      Result := True;
    end else
    begin
      PGIBox('Pas d''affaire à importer');
    end;
  end;
end;

function T_SPIGAOCONNECT.ChargeAO(CodeAffaire, CodeTiers,CodeEtab: String; IdSPIGAO : double): Boolean;
Var Projet : IProject;
    req : String;
begin
  result := False;
  if not fconnected  then Exit;

  ArticleDivers := GetParamSoc ('SO_BTARTICLEDIV');
  if ArticleDivers = '' then BEGIN PGIBox (TraduireMemoire('Article divers non défini'),'Import DIE'); exit; END;
  ArticleDivers := CodeArticleUnique(ArticleDivers,'','','','','');

  Projet := ChargeProjet(False, '', IdSPIGAO);
  if Projet <> nil then
  begin
    // Création TOB (entete piece)
    TOBPiece := Tob.Create ('PIECE', nil, -1);
    TRY
      TobPiece.AddChampSupValeur ('NATUREPIECEG', 'ETU');
      TobPiece.AddChampSupValeur ('AFFAIRE', CodeAffaire);
      TobPiece.AddChampSupValeur ('TIERS', CodeTiers);
      TobPiece.AddChampSupValeur ('ETABLISSEMENT', CodeEtab);
      TobPiece.AddChampSupValeur ('DOMAINE', '');
      TobPiece.AddChampSupValeur ('DATEPIECE', V_PGI.DateEntree);
      TobPiece.AddChampSupValeur ('REFINTERNE', projet.BillOfQuantity.Title);

      if projet.BillOfQuantity.WorkItems.Count > 0 then
      begin
        ChargeTob (projet.BillOfQuantity.WorkItems);
        Result := CreatePieceFromTOB (TOBPiece,nil,nil,nil);
        if Result then
        begin
          Req := 'SELECT BDE_AFFAIRE FROM BDETETUDE WHERE '+
                  'BDE_AFFAIRE="'+CodeAffaire+'" AND BDE_ORDRE=1 AND '+
                  'BDE_NATUREDOC="002" AND BDE_INDICE=0"';
          if not ExisteSql (req) then
          begin
            Req := 'INSERT INTO BDETETUDE '+
              '(BDE_AFFAIRE,BDE_ORDRE,BDE_NATUREDOC,BDE_INDICE,BDE_TYPE,BDE_QUALIFNAT,BDE_DESIGNATION,'+
              'BDE_NAME,BDE_TIERS,BDE_PIECEASSOCIEE,'+
              'BDE_SELECTIONNE,BDE_DATEDEPART,'+
              'BDE_DATEFIN,BDE_NATUREAUXI,BDE_CLIENT,'+
              'BDE_DATECREATION,BDE_DATEMODIF,'+
              'BDE_NATUREPIECEG,BDE_SOUCHE,'+
              'BDE_NUMERO,BDE_INDICEG,BDE_FOURNISSEUR)'+
              ' VALUES '+
              '("'+ CodeAffaire +'",1,"002",0,"003","PRINC","INTEGRATION SPIGAO","';
            Req := Req + projet.Path + '","","' + EncodeRefPiece(TOBPIECE);
            req := req + '","X","' + usdatetime(StrToDate(TOBPiece.GetValue('GP_DATEPIECE'))) +'","';
            req := req + usdatetime(iDate2099) + '","CLI","' + CodeTiers + '","';
            req := req + usdatetime(Now) +'","'+ usdatetime(Now) +'","';
            req := req + TOBPiece.GetValue('GP_NATUREPIECEG') + '","' + TOBPiece.GetValue('GP_SOUCHE')+ '",';
            req := req + IntToStr(TOBPiece.GetValue('GP_NUMERO'))+ ',0,"")';
          end else
          begin
            req := 'UPDATE BDETETUDE SET '+
                   'BDE_DATEDEPART="'+usdatetime(StrToDate(TOBPiece.GetValue('GP_DATEPIECE')))+'",'+
                   'BDE_DATECREATION="'+usdatetime(Now)+'",BDE_DATEMODIF="'+usdatetime(Now)+'",'+
                   'BDE_NATUREPIECEG="'+TOBPiece.GetValue('GP_NATUREPIECEG')+'",'+
                   'BDE_SOUCHE="'+TOBPiece.GetValue('GP_SOUCHE')+'",'+
                   'BDE_NUMERO='+IntToStr(TOBPiece.GetValue('GP_NUMERO'));
          end;
          ExecuteSQL(Req);
          if fNewgestion then fDetect.Deals.MarkDealAsIntegrated(StrToInt(FloatToStr(IdSPIGAO)),true);
        end;
      end;
    FINALLY
      TOBPiece.free;
      TOBPiece := nil;
    END;
  end;
end;

procedure T_SPIGAOCONNECT.ChargeNewAffaires;
var retour : string;
		StatutAffaire : string;
begin
  V_PGI.ZoomOLE := true;  //Affichage en mode modal
  Retour := AGLLanceFiche ('BTP','BTAFFAIRE','','','ACTION=CREATION;STATUT:PRO;AFF_TIERS:;ETAT:SPIGAO;NATURE:ETU;SPIGAONEW');
  V_PGI.ZoomOLE := false;  //Affichage en mode modal
  If Retour <> '' then
  begin
		V_PGI.ZoomOLE := true;  //Affichage en mode modal
    StatutAffaire := ReadTokenSt(Retour);
    AGLLanceFiche('BTP','BTAFFAIRE','',Retour,'ACTION=MODIFICATION;STATUT:'+Statutaffaire);
		V_PGI.ZoomOLE := false;  //Affichage en mode modal
{$IFNDEF PGIMAJVER}
{$IFNDEF MOBILITE}
{$IFNDEF POUCHAIN}
{$IFNDEF SPINNAKER}
    AfterChangeModule(145); // pour recharger les nombres d'affaires dans les libellés du menu
{$ENDIF}
{$ENDIF}
{$ENDIF}
{$ENDIF}
  end;
end;

function T_SPIGAOCONNECT.ChargeProjet(Direct: Boolean; CodeAffaire: String; IdSpigao : double = 0): IProject;
var Import : IImportResult;
    Filepath : string;
    QQ: TQuery;
begin
  Result := nil;
  if Direct then
  begin
    QQ := OpenSql ('SELECT BDE_NAME FROM BDETETUDE WHERE BDE_AFFAIRE="'+CodeAffaire+'"',true,-1,'',true);
    if Not QQ.Eof then
    begin
      Filepath := QQ.Fields [0] .AsString;
      if FileExists(Filepath) then Result := fLoader.Processing.LoadProject(Filepath);
    end else
    begin
      PGIBox('Projet introuvable pour cette affaire. Traitement impossible.');
    end;
    ferme (QQ);
  end else
  begin
    if fNewgestion then
    begin
      ControleAcces;
			Import := fLoader.Processing.UI.ImportProject(StrToInt(FloatToStr(IdSpiGao)));
    end else
    begin
      Import := fLoader.processing.ui.ShowImportWizard(nil, False);
    end;
    if Import <> nil then Result := Import.project;
  end;
end;

procedure T_SPIGAOCONNECT.ChargeTob(currentCollection: IWorkItemCollection;Niv: Integer);
var i, niveau : Integer;
    workItem : IWorkItem;
    Dim : string;
begin
  For i := 0 To currentCollection.count - 1 do
  begin
    workItem := currentCollection.item[i];
    // Création TOB pour ligne de pièce
    TobLignes := Tob.Create ('', TobPiece, -1);
    TobLignes.AddChampSupValeur ('TYPELIGNE', 'COM');
    TobLignes.AddChampSupValeur ('NIVEAUIMBRIC', Niv);
    TobLignes.AddChampSupValeur ('LIBELLE', workItem.designation.Current.value);
    TobLignes.AddChampSupValeur ('IDSPIGAO', workItem.Id);
    if workItem.AsPrice = nil then
    begin
      TobLignes.AddChampSupValeur ('CODEARTICLE', '');
      TobLignes.AddChampSupValeur ('ARTICLE', '');
      TobLignes.AddChampSupValeur ('REFARTSAISIE', workItem.PriceNumber.Current.value);
      TobLignes.AddChampSupValeur ('NUMEROTATION', workItem.PriceNumber.Current.value);
      TobLignes.AddChampSupValeur ('NUMFORCED', 'X');
      TobLignes.AddChampSupValeur ('BLOCNOTE', '');
      TobLignes.AddChampSupValeur ('QTEFACT', 0);
      TobLignes.AddChampSupValeur ('QUALIFQTE', '');
      TobLignes.AddChampSupValeur ('PUHTDEV', 0);
      TobLignes.AddChampSupValeur ('DATELIVRAISON', V_PGI.DateEntree);
      TobLignes.AddChampSupValeur ('DEPOT', '');
      TobLignes.AddChampSupValeur ('BLOQUETARIF', '-');
    end else
    begin
      TobLignes.PutValue ('TYPELIGNE', 'ART');
      TobLignes.AddChampSupValeur ('CODEARTICLE', trim(CodeArticleGenerique(ArticleDivers,dim,dim,dim,dim,dim)));
      TobLignes.AddChampSupValeur ('ARTICLE', ArticleDivers);
      TobLignes.AddChampSupValeur ('REFARTSAISIE', workItem.PriceNumber.Current.value);
      TobLignes.AddChampSupValeur ('NUMEROTATION', workItem.PriceNumber.Current.value);
      TobLignes.AddChampSupValeur ('NUMFORCED', 'X');
      TobLignes.AddChampSupValeur ('BLOCNOTE', workItem.designation.Current.value);
      TobLignes.AddChampSupValeur ('QTEFACT', workItem.AsPrice.Quantity.Current.value);
      TobLignes.AddChampSupValeur ('QUALIFQTE', workItem.AsPrice.Unit_.Current.Value);
      TobLignes.AddChampSupValeur ('PUHTDEV', workItem.AsPrice.UnitPrice.Current.value);
      TobLignes.AddChampSupValeur ('DATELIVRAISON', V_PGI.DateEntree);
      TobLignes.AddChampSupValeur ('DEPOT', '');
      If workItem.AsPrice.IsFixed then TobLignes.AddChampSupValeur ('BLOQUETARIF', 'X')
                                  else TobLignes.AddChampSupValeur ('BLOQUETARIF', '-');
    end;
    if workItem.AsChapter <> nil then
    begin
      niveau := workItem.Level.Current.Value;
      TobLignes.AddChampSupValeur ('TYPELIGNE', 'DP'+ IntToStr(niveau));
      TobLignes.AddChampSupValeur ('NIVEAUIMBRIC', Niveau);
      ChargeTob(workItem.AsChapter.WorkItems, niveau);
      TobLignes := Tob.Create ('', TobPiece, -1);
      TobLignes.AddChampSupValeur ('TYPELIGNE', 'TP'+ IntToStr(niveau));
      TobLignes.AddChampSupValeur ('NIVEAUIMBRIC', Niveau);
      TobLignes.AddChampSupValeur ('LIBELLE', 'Total : '+ workItem.designation.Current.value);
      TobLignes.AddChampSupValeur ('CODEARTICLE', '');
      TobLignes.AddChampSupValeur ('ARTICLE', '');
      TobLignes.AddChampSupValeur ('REFARTSAISIE', workItem.PriceNumber.Current.value);
      TobLignes.AddChampSupValeur ('NUMEROTATION', workItem.PriceNumber.Current.value);
      TobLignes.AddChampSupValeur ('NUMFORCED', 'X');
      TobLignes.AddChampSupValeur ('QTEFACT', 0);
      TobLignes.AddChampSupValeur ('QUALIFQTE', '');
      TobLignes.AddChampSupValeur ('PUHTDEV', 0);
      TobLignes.AddChampSupValeur ('DATELIVRAISON', V_PGI.DateEntree);
      TobLignes.AddChampSupValeur ('DEPOT', '');
    end;
  end;
end;

procedure T_SPIGAOCONNECT.Confdetection;
begin
  if not fconnected  then Exit;
  fLoader.detection.ui.ShowUserSettings;
{$IFNDEF PGIMAJVER}
{$IFNDEF MOBILITE}
{$IFNDEF POUCHAIN}
{$IFNDEF SPINNAKER}
	if fDetect.Account.type_ <> AccountType_Undefined then AfterChangeModule(145);
{$ENDIF}
{$ENDIF}
{$ENDIF}
{$ENDIF}
end;

procedure T_SPIGAOCONNECT.ConfTraitement;
begin
  if not fconnected  then Exit;
  fLoader.processing.ui.ShowUserSettings;
end;

function T_SPIGAOCONNECT.Connecte: Boolean;
var Path,S : string;
BEGIN
  fconnected := False;
  S := GetFromRegistry(HKEY_LOCAL_MACHINE,'SOFTWARE\Edisys\IULM\','Path','',TRUE);
  if S <> '' then
  begin
  	Path := IncludeTrailingBackslash(S) + 'Edisys.IULM.Core.dll';
  end else
  begin
  Path := IncludeTrailingBackslash(TCBPPath.GetProgramFiles) + 'Edisys\IULM\Edisys.IULM.Core.dll';
  end;
  if FileExists(Path) then
  begin
    try
      fLoader := CoSpigaoLoader.Create;
      fdetect := fLoader.detection;
      fVersion := fLoader.Version.Text;
      if (fVersion > '2.2')  and (fversion < '3.2') then
      begin
      	fLoader.Unlock('LSE', 'LSE Business BTP', '502517D585D5F02CE657');
        fOldGestion := True;
        PGIInfo ('ATTENTION : Une nouvelle version de SPIGAO existe#13#10 Veuillez vous mettre à jour.');
        fconnected := True;
      end else if (fVersion >= '3.2') then
      begin
        //
        fXX := TEdisys_IULM_CoreIDetectionEvents2.create (Application.mainform);
        fXX.Connect(fDetect.Events);
      	fLoader.Unlock('LSE', 'LSE Business BTP', '502517D585D5F02CE657');
        fconnected := True;
        fNewgestion := True;
        SetGlobalSettings;
        SetEvents;
      end else
      begin
        PgiError ('ATTENTION : La version de SPIGAO ne peut être utilisée.#13#10 Merci de vous mettre à jour');
        fLoader := Nil;
        fdetect := Nil;
      end;
    except
      PGIError('Version de SPIGAO non conforme.');
    end;
  end;
  Result := fconnected;
end;

procedure T_SPIGAOCONNECT.ControleAcces;
begin
	if fDetect.Account.type_ = AccountType_Undefined then
  begin
  	fLoader.detection.ui.ShowUserSettings;
  end;
end;

constructor T_SPIGAOCONNECT.create;
begin
  TOBDETETUDE := TOB.Create ('LES MEMO DETETUDE',nil,-1);
  fconnected := false;
  fVersion := '';
end;

procedure T_SPIGAOCONNECT.deconnecte;
begin
  if not fconnected then exit;
	if fNewgestion Then fXX.Free;
  fLoader := Nil;
  fdetect := Nil;
end;

destructor T_SPIGAOCONNECT.destroy;
begin
  TOBDETETUDE.free; 
  deconnecte;
  inherited;
end;

function T_SPIGAOCONNECT.ExportAO(CodeAffaire: String): Boolean;
Var Projet : IProject;
    workItem : IWorkItem;
    workItemCollec : IWorkItemCollection;
    i : Integer;
    QQ: TQuery;
    Retour : ExecResult;
begin
  result := False;
  if not fconnected  then Exit;

  Projet := ChargeProjet( True, CodeAffaire);
  if Projet <> nil then
  begin
    workItemCollec := Projet.BillOfQuantity.GetFlatWorkItemList(ComWorkItemFilter_PriceAll);
    For i := 0 to workItemCollec.count-1 do
    begin
      workItem := workItemCollec.item[i];

      // Lecture ligne d'étude correspondant à l'idenfiant de ligne
      QQ := OpenSql ('SELECT GL_QTEFACT, GL_PUHTDEV FROM LIGNE WHERE GL_NATUREPIECEG="ETU" AND GL_AFFAIRE="'+CodeAffaire+'" AND GL_IDSPIGAO="'+workItem.Id+'"',true,-1,'',true);
      if Not QQ.Eof then
      begin
        workItem.AsPrice.SetQuantity(QQ.Fields [0] .AsFloat);
        workItem.AsPrice.SetUnitPrice(QQ.Fields [1] .AsFloat);
      end;
      ferme (QQ);
    end;
    if workItemCollec.count > 0 then
    begin
      Retour := fLoader.Processing.UI.ShowExportWizard(Projet, '');
    end;
  end;
end;

procedure T_SPIGAOCONNECT.Infos;
begin
  PGIBox('Le logiciel SPIGAO n''est pas installé sur votre système.'+chr(13)+chr(13)+' Veuillez contacter notre service commercial. Merci.','Fonction non disponible');
end;

procedure T_SPIGAOCONNECT.IntegreDie ;
var retour : string;
		StatutAffaire : string;
begin
  V_PGI.ZoomOLE := true;  //Affichage en mode modal
  Retour := AGLLanceFiche ('BTP','BTAFFAIRE','','','ACTION=CREATION;STATUT:PRO;AFF_TIERS:;ETAT:SPIGAO;NATURE:ETU');
  V_PGI.ZoomOLE := false;  //Affichage en mode modal
  If Retour <> '' then
  begin
		V_PGI.ZoomOLE := true;  //Affichage en mode modal
    StatutAffaire := ReadTokenSt(Retour);
    AGLLanceFiche('BTP','BTAFFAIRE','',Retour,'ACTION=MODIFICATION;STATUT:'+Statutaffaire);
		V_PGI.ZoomOLE := false;  //Affichage en mode modal
{$IFNDEF PGIMAJVER}
{$IFNDEF MOBILITE}
{$IFNDEF POUCHAIN}
{$IFNDEF SPINNAKER}
  	AfterChangeModule (145);
{$ENDIF}
{$ENDIF}
{$ENDIF}
{$ENDIF}
  end;
end;

procedure T_SPIGAOCONNECT.LienBTP;
begin
  if not fconnected  then Exit;
  LanceWeb(fdetect.WebSite.HomePageUrl, False);
end;

procedure T_SPIGAOCONNECT.MajStatusAffaire(MonAffaireId: Integer);
var dc : IDeal;
begin
  if not fconnected  then Exit;
  dc := fDetect.LoadDeal(MonAffaireId);
  fDetect.SetIntegrationStatus(dc.Id, True);
end;

procedure T_SPIGAOCONNECT.MonteToaster;
begin
  // modif BRL 7/11 suite pb tps de réponse chez CFF
  Exit ;

	if Not fconnected then exit;
  if not fNewgestion then Exit;
	fLoader.Detection.UI.Notifications.ShowAlert;
end;

procedure T_SPIGAOCONNECT.RecupNbAffaires(var Nbnew, NbMod, NbImp,NbMes: Integer);
var Re1 : ISearchResult;
    Re2 : IDealSummaryCollection;
    Re3 : ISpigaoMessageCollection;
    nb : integer;
begin
  // modif BRL 7/11 suite pb tps de réponse chez CFF
  Exit ;
  
  if not fconnected  then Exit;
  Re1 := fdetect.SearchNewDeals;
  Nbnew := Re1.Count;

  Re1 := fdetect.SearchUpdatedDeals;
  NbMod := Re1.Count;

  Re2 := fdetect.ListDealsToLoad;
  NbImp := Re2.FullCount;

  Re3 := fdetect.ListSpigaoMessages;
  NbMes := Re3.FullCount;
end;


procedure T_SPIGAOCONNECT.SetEvents;
begin
	fXX.UserRequestNewDeals := UserRequestNewDeals;
  fXX.UserRequestUpdatedDeals := UserRequestUpdatedDeals;
  fXX.UserRequestDealsWithAvailableDIE := UserRequestDealsWithAvailableDIE;
end;

procedure T_SPIGAOCONNECT.SetGlobalSettings;
begin
  with fLoader.Processing.GlobalSettings do
  begin
    ChapterImportMode := ImportMode_Allow;
    ChapterMandatory := false;
    CommentImportMode  := ImportMode_Allow;
    DesignationMaxLength  := 70;
    IdentAllowAlpha := True;
    IdentMaxLength := 0;
    MaxLevel := 9;
		MaxLevelAction:=MaxLevelAction_Bottom;
		MaxLineCount :=0;
		PriceNumberAllowAlpha := true;
    PriceNumberForComment :='';
    PriceNumberMandatory:=false;
    PriceNumberMaxLength :=35;
    PriceNumberRenumberingScope := PriceNumberRenumberingScope_Global;
    QuantityDecimals :=3;
    QuantityDefaultValue :=0.0;
    QuantityMandatory :=True;
    UnitAllowedChars := '';
    UnitDefaultValue := 'U';
    UnitMaxLength :=3;
    UnitPriceAllowNegative :=True;
    UnitPriceDecimals :=2;
    UnitReplaceString :='*';
  end;
end;

function T_SPIGAOCONNECT.TelechargerDIE(Id: Integer): Boolean;
var Lien : string;
begin
  Result := False;
  if not fconnected  then Exit;

  Lien := fDetect.WebSite.GetDIEPageUrl(Id);
  LanceWeb(Lien,False);
end;

function T_SPIGAOCONNECT.TestAffaire(CodeAffaire: String): Boolean;
Var Q : TQuery ;
    Req : String;
    TOBD : TOB;
BEGIN
  Result := False;
  if not fconnected  then Exit;
  if CodeAffaire = '' then exit;

  Result := (StrToInt(GetChampsAffaire (CodeAffaire,'AFF_IDSPIGAO'))<>0);

  if Result = False then
  begin
  // Traitement au cas où l'intégration SPIGAO a été faite sur une affaire déjà créée
  // sans passer par le chargement depuis les affaires à importer
  // Dans ce cas, le champ AFF_IDSPIGAO n'est pas renseigné, il faut donc aller vérifier le type de doucment dans BDETETUDE
    TOBD := TOBDETETUDE.findFirst(['AFFAIRE'],[Codeaffaire],true);
    if TOBD = nil then
    begin
      TOBD := TOB.Create ('UNE AFFAIRE',TOBDETETUDE,-1);
      TOBD.AddChampSupValeur('AFFAIRE',Codeaffaire);
      TOBD.AddChampSupValeur('BDE_TYPE','');
      Req:='SELECT BDE_TYPE FROM BDETETUDE WHERE BDE_AFFAIRE="'+Codeaffaire+'"' ;
      Q:=OpenSQL(Req,TRUE,-1,'',true) ;
      if Not Q.EOF then
      begin
        TOBD.SetString('BDE_TYPE',Q.Fields[0].AsString);
        if Q.Fields[0].AsString = '003' then Result := True;
      end;
      Ferme(Q) ;
    end else
    begin
      if TOBD.GetString('BDE_TYPE')='003' then Result := True;
    end;
  end;
end;

procedure T_SPIGAOCONNECT.UserRequestDealsWithAvailableDIE(Sender : TObject);
begin
	IntegreDie;
end;

procedure T_SPIGAOCONNECT.UserRequestNewDeals(Sender: TObject);
begin
  if fNewgestion then ChargeNewAffaires;
end;

procedure T_SPIGAOCONNECT.UserRequestUpdatedDeals(Sender: TObject);
begin
	VoirAffairesMod;
end;

function T_SPIGAOCONNECT.VoirAffaire(Id: Integer): Boolean;
var Lien : string;
begin
  Result := False;
  if not fconnected  then Exit;
  if fNewgestion then
  begin
  	ControleAcces;
    fdetect.UI.Deals.ViewDeal(Id);
  end else
  begin
    Lien := fDetect.WebSite.GetDealPageUrl(Id);
    LanceWeb(Lien,False);
  end;
end;

procedure T_SPIGAOCONNECT.VoirAffairesMod ;
var Re1 : ISearchResult;
begin
  if not fconnected  then Exit;
  if fNewgestion then
  begin
  	ControleAcces;
		fDetect.UI.Deals.ListUpdatedDeals;
  end else
  begin
  	Re1 := fdetect.SearchUpdatedDeals;
	  LanceWeb(Re1.Url,False);
  end;
end;

procedure T_SPIGAOCONNECT.VoirAffairesNew;
var Re1 : ISearchResult;
		Omg : IDealResult;
begin
  if not fconnected  then Exit;
  if fNewgestion then
  begin
  	ControleAcces;
    Omg := fDetect.UI.Deals.ListNewDeals;
  end else
  begin
    Re1 := fdetect.SearchNewDeals;
    LanceWeb(Re1.Url,False);
  end;
end;

function T_SPIGAOCONNECT.VoirBPX(CodeAffaire,Identifiant: string): Boolean;
Var Projet : IProject;
begin
  Result := False;
  if not fconnected  then Exit;
  Projet := ChargeProjet(True, CodeAffaire);
  if Projet <> nil then
  begin
    if fNewgestion then
    begin
    	ControleAcces;
      fLoader.Processing.UI.ViewDocuments(Projet,Identifiant,ComDocumentType_BPX)
    end else
    begin
    	fLoader.processing.ui.ShowBPX(projet, Identifiant, ComDocumentType_BPX);
    end;
  end;
end;

procedure T_SPIGAOCONNECT.VoirMessages;
begin
  if not fconnected  then Exit;
  if fNewgestion then
  begin
  	ControleAcces;
    fDetect.UI.Notifications.ShowMessages;
  end else
  begin
  	LanceWeb(fdetect.WebSite.HomePageUrl, False);
  end;
end;

End.


