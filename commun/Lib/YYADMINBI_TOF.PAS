{**********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 10/10/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : YYADMINBI ()
Mots clefs ... : TOF;YYADMINBI
*****************************************************************}
Unit YYADMINBI_TOF ;

Interface

Uses
{$IFDEF VER150}
  variants,
{$ENDIF}
     Wintypes,
     StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     OPENOLAP4Lib_TLB,
     OPENOLAPScriptLib_TLB,
     ED_TOOLS,
     HTb97,
     uTob,
     Registry
     ;

Type
  TOF_YYADMINBI = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private

    Action:String;
    ChkAnalD,ChkAnonymous: TCheckBox;
    EServeurCwas, ELoginCwas, EPasswdCwas,EDomain,EServeurBi,ECodeCube, ELibcube: THEdit;
    CDatamart,CCube:THValCombobox ;
    BtnTrtDatamart,BtnAddPersos,BtnDelPersos,BtnTestConnect:TToolbarButton97;
    TVDimensions,TVPersos:TTreeView;
    TDim,TPerso:TOB;
    procedure ChkAnalDOnClick (Sender:Tobject);
    procedure ChkAnonymousOnClick (Sender:Tobject);
    procedure BtnTrtDatamartOnClick (Sender:Tobject);
    procedure BtnAddPersosOnClick (Sender:Tobject);
    procedure BtnDelPersosOnClick (Sender:Tobject);
    procedure BtnTestConnectOnClick (Sender:Tobject);
    procedure EServeurbiOnExit (Sender:Tobject);
    procedure CDatamartOnExit (Sender:Tobject);
    procedure CCubeOnExit (Sender:Tobject);
    procedure ECodeCubeOnExit (Sender:Tobject);
    procedure CDatamartOnClick (Sender:Tobject);
    procedure ChargeDimensions;
    function ActiveOpenOlap:boolean;
    function DsoAdd(dim,table,perso,champlibrechoisi,champlibelle,tabletteval:string):boolean;
    function DsoDel(dim,perso:string):boolean;
    procedure RemplirCDatamart;
    procedure RemplirCCube;
    function okConnectDatamart:olevariant;
    function okConnectSrv:olevariant;
    function okConnectCube:olevariant;
  end ;

type monObjet = Class(TObject)
  private
    FOwner:TComponent;
    FCodeCube:string;
    procedure SetCodeCube(const Value: string);
    procedure SetLibCube(const Value: string);
    procedure SetactiveBi(const Value: boolean);
    procedure Setanonymous(const Value: boolean);
    procedure SetserveurBi(const Value: string);
    procedure SetDataMart(const Value: string);
    procedure SetCube(const Value: string);
    procedure SetServeurCwas(const Value: string);
    procedure SetLogin(const Value: string);
    procedure SetPassword(const Value: string);
    procedure SetDomaine(const Value: string);
    procedure SetGuidBase(const Value: string);
    function GetCodeCube: string;
    function GetLibCube: string;
    function GetServeurBI: string;
    function GetactiveBI: boolean;
    function Getanonymous: boolean;
    function GetDatamart: string;
    function GetCube: string;
    function GetServeurCwas: string;
    function GetLogin: string;
    function GetPassword: string;
    function GetDomaine: string;
    function GetGuidBase: string;
  public
    constructor create(Aowner:tComponent);
    destructor destroy();override;
  published
    property Anonymous: boolean  read Getanonymous write Setanonymous;
    property CodeCube: string  read GetCodeCube write SetCodeCube;
    property LibCube: string  read GetLibCube write SetLibCube;
    property ServeurBi: string  read GetserveurBi write SetserveurBi;
    property ActiveBi: boolean  read GetactiveBi write SetactiveBi;
    property GuidBase: string  read GetGuidBase write SetGuidBase;
    property DataMart: string  read GetDataMart write SetDataMart;
    property Cube: string  read GetCube write SetCube;
    property ServeurCwas: string  read GetServeurCwas write SetServeurCwas;
    property Login: string  read GetLogin write SetLogin;
    property Password: string  read GetPassword write SetPassword;
    property Domaine: string  read GetDomaine write SetDomaine;
  end;

Var myTOB,myTOBFille,TRefCubeOlap:TOB;
    v_connolap:monobjet;

Implementation

uses paramsoc,utobdebug;

const
    FSeriaFree = 'fLVs-VvlJ-F7uv-7D0E'; //permet d'utiliser les activeX sans seria

procedure TOF_YYADMINBI.OnArgument (S : String ) ;
var CleActiveOpenOlap,ArgCodeCube,stwhere:String;
    Qmn:TQuery;
//    Key:string;
//    Reg:TRegistry;
begin
  Inherited ;

  ChkAnalD:=TCheckBox(GetControl('CHKANALD'));
  ChkAnonymous:=TCheckBox(GetControl('CBANONYMOUS'));  

  CDatamart:=THValComboBox(GetControl('CDATAMART'));
  CCube:=THValComboBox(GetControl('CCUBE'));

  ECodeCube:=THEdit(GetControl('ECODECUBE'));

  EServeurBi:=THEdit(GetControl('ESERVEURBI'));
  ELibCube:=THEdit(GetControl('ELIBCUBE'));

  EServeurCwas:=THEdit(GetControl('ESERVEURCWAS'));
  ELoginCwas:=THEdit(GetControl('ELOGINCWAS'));
  EPasswdCwas:=THEdit(GetControl('EPASSWDCWAS'));
  EDomain:=THEdit(GetControl('EDOMAIN'));

  BtnTrtDatamart:=TToolbarButton97(GetControl('BTNTRTDATAMART'));

  BtnAddPersos:=TToolbarButton97(GetControl('BTNADDPERSO'));
  BtnDelPersos:=TToolbarButton97(GetControl('BTNDELPERSO'));

  BtnTestConnect:=TToolbarButton97(GetControl('BTNTESTCONNECT'));

  TVDimensions:=TTreeView(GetControl('TVDIMENSIONS'));
  TVPersos:=TTreeView(GetControl('TVPERSOS'));

  ChkAnalD.OnClick:=ChkAnalDOnClick;
  ChkAnonymous.OnClick:=ChkAnonymousOnClick;

  ECodeCube.OnExit := ECodeCubeOnExit;

  EServeurbi.OnExit := EServeurBiOnExit;

  BtnTrtDatamart.OnClick:=BtnTrtDatamartOnClick;
  BtnAddPersos.OnClick:=BtnAddPersosOnClick;
  BtnDelPersos.OnClick:=BtnDelPersosOnClick;
  BtnTestConnect.OnClick:=BtnTestConnectOnClick;

  CDatamart.onExit:=cDatamartOnExit;
  CCube.onExit:=cCubeOnExit;

  CDatamart.onClick:=cDatamartOnClick;

  ArgCodeCube:=ReadTokenSt(S);
  Action:=ReadTokenSt(S);

  myTOB:=TOB.Create('YCUBEOLAP',nil,-1);
  if (ArgCodeCube <> '') and (Action='ACTION=MODIFICATION')
    then myTOB.LoadDetailDB('YCUBEOLAP','"'+ArgCodeCube+'"','',nil,False)
    else myTOBFille := TOB.Create('YCUBEOLAP',myTOB,-1);

  v_connolap:=monobjet.create(appli);
  if not assigned(v_connolap) then exit;

  SetControlText('ECodeCube',v_connolap.codecube);
  ChkAnalD.Checked:=v_connolap.ActiveBi;
  ChkAnonymous.Checked:=v_connolap.anonymous;

  if Action='ACTION=CREATION' then
  begin
    ECodeCube.Enabled:=True;
    ChkAnonymous.Checked:=True;
    ChkAnonymousOnclick(nil);
  end
  else
  begin
    BtnTrtDatamart.Hint:= BtnTrtDatamart.Hint + ' ' + v_connolap.datamart;
    ECodeCube.Enabled:=False;
    CleActiveOpenOlap:='';

  {  Reg := TRegistry.Create;
    Key:='SOFTWARE\Open Executive 4\OpenOLAP';
    try
      Reg.RootKey := HKEY_LOCAL_MACHINE;
      if not Reg.KeyExists(Key) then
      begin
        CleActiveOpenOlap:=GetFromRegistry(Reg.RootKey,Key ,'PullBufferSize','');
        Reg.Free;
        Exit;
      end;
    finally
      Reg.Free;
    end;}

    If (CleActiveOpenOlap='') then CleActiveOpenOlap:=booltostr(v_connolap.ActiveBi); //je ne parviens pas a taper dans la registry
    if CleActiveOpenOlap <> '' then
    begin
      ChkAnalD.Checked:=True;
      ChkAnalD.Enabled:=False;
      ActiveOpenOlap;
    end
    else ChkAnalD.Enabled:=True;

    SetControlText('EServeurBi',v_connolap.serveurbi);
    CDatamart.Text:=v_connolap.datamart;
    CCube.Text:=v_connolap.cube;
    SetControlText('ELibCube',v_connolap.libcube);
    SetControlText('EServeurCwas',v_connolap.serveurcwas);
    SetControlText('ELoginCwas',v_connolap.login);
    SetControlText('EPasswdCwas',v_connolap.password);
    SetControlText('EDomain',v_connolap.domaine);

    if (v_connolap.CodeCube <> '') then
    begin
      TRefCubeOlap:=TOB.Create('Referentiel',nil,-1);
      TRefCubeOlap.AddChampSup('YRU_CHAMPLIBELLE',true);
      TRefCubeOlap.AddChampSup('YRU_TABLETTE',true);
      stwhere := 'WHERE YRU_YBOGUID="'+v_connolap.guidbase+'"';
      try
        QMn :=  OpenSQL('SELECT * FROM YREFCUBEOLAP ' + stWhere,True);
        try
          if not QMn.Eof then TRefCubeOlap.LoadDetailDB('YREFCUBEOLAP','','',QMn,true);
        finally
          Ferme(QMn);
        end;
      finally
      end;
      CDatamart.Enabled:=false;
      CCube.Enabled:=false;

      ChargeDimensions;
    end;
  end;
end;

procedure TOF_YYADMINBI.OnUpdate ;
var iInd:integer;
    cubemodifie:boolean;
    QMn:TQuery;
    TOBBaseOlap,TOBBaseOlapDet:TOB;
    stWhere:string;
begin
  if getControlText('ECodeCube')='' then
  begin
    PGIBOX('Saisie du code cube obligatoire !');
    SetFocusControl('ECodeCube');
    exit;
  end;

  Inherited ;

  cubemodifie:=false;

  if not assigned(v_connolap) then exit;

  v_connolap.codecube:=GetControlText('ECodeCube');
  v_connolap.serveurbi:=GetControlText('EServeurBi');
  v_connolap.datamart:=CDatamart.Text;
  v_connolap.cube:=CCube.Text;
  v_connolap.libcube:=GetControlText('ELibCube');
  v_connolap.serveurcwas:=GetControlText('EServeurCwas');
  v_connolap.login:=GetControlText('ELoginCwas');
  v_connolap.password:=GetControlText('EPasswdCwas');
  v_connolap.domaine:=GetControlText('EDomain');


  if not assigned(myTOB) then exit;
  myTOB.Detail[0].InsertOrUpdateDb;

  TOBBaseOlap:=TOB.Create('Ma Base',nil,-1);

  stwhere := 'WHERE YBO_DATAMART="' +CDatamart.Text+'"';
  try
  QMn :=  OpenSQL('SELECT * FROM YBASEOLAP ' + stWhere,True);
    try
      if not QMn.Eof then
        begin
          TOBBaseOlap.LoadDetailDB('YBASEOLAP','','',QMn,true);
          v_connolap.GuidBase:=TOBBaseOlap.Detail[0].getString('YBO_GUID');
          myTOB.Detail[0].InsertOrUpdateDb;
        end
        else
        begin
          v_connolap.GuidBase:=AglGetGuid();
          myTOB.Detail[0].InsertOrUpdateDb;
          TOBBaseOlapDet:=TOB.Create('YBASEOLAP',TOBBaseOlap,-1);
//          TOBBaseOlap.Detail[0].AddChampSupValeur('YBO_GUID',v_connolap.guidbase);
//          TOBBaseOlap.Detail[0].AddChampSupValeur('YBO_DATAMART',v_connolap.datamart);
          TOBBaseOlap.Detail[0].PutValue('YBO_GUID',v_connolap.guidbase);
          TOBBaseOlap.Detail[0].PutValue('YBO_DATAMART',v_connolap.datamart);
         end
    finally
      Ferme(QMn);
    end;
  finally
    if TOBBaseOlap.Detail.Count>0
      then TOBBaseOlap.Detail[0].InsertOrUpdateDB;
    FreeAndNil(TOBBaseOlap);
  end;


  if not assigned (TrefCubeOlap) then exit;
  if TRefCubeOlap.FillesCount(100)=0 then exit;

  for iInd:=0 to TRefCubeOlap.Detail.Count - 1 do
  begin
    if TRefCubeOlap.Detail[iInd].IsFieldModified('YRU_ETAT') then
      begin
        if TRefCubeOlap.Detail[iInd].getValue('YRU_ETAT') = 'X'
          then
            begin
              if not DsoAdd(
              TRefCubeOlap.Detail[iInd].getValue('YRU_DIMENSION'),
              TRefCubeOlap.Detail[iInd].getValue('YRU_TABLEOLAP'),
              TRefCubeOlap.Detail[iInd].getValue('YRU_ATTRIBUT'),
              TRefCubeOlap.Detail[iInd].getValue('YRU_SOURCEKEY'),
              TRefCubeOlap.Detail[iInd].getValue('YRU_CHAMPLIBELLE'),
              TRefCubeOlap.Detail[iInd].getValue('YRU_TABLETTE'))
                then PGIBOX('Erreur lors de la personnalisation avec le champ : '
                  + TRefCubeOlap.Detail[iInd].getValue('YRU_ATTRIBUT'))
                else cubemodifie:=true;
            end
          else if not DsoDel(
            TRefCubeOlap.Detail[iInd].getValue('YRU_DIMENSION'),
            TRefCubeOlap.Detail[iInd].getValue('YRU_ATTRIBUT'))
              then PGIBOX('Erreur lors de la personnalisation avec le champ : '
                + TRefCubeOlap.Detail[iInd].getValue('YRU_ATTRIBUT'))
              else cubemodifie:=true;
      end;
  end;
  if cubemodifie then
    begin
      BtnTrtDatamartOnClick(nil);
      TRefCubeOlap.InsertOrUpdateDB;
    end;
end ;


procedure TOF_YYADMINBI.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_YYADMINBI.OnDelete ;
begin
  Inherited ;
  if assigned(myTOB) then myTOB.Detail[0].DeleteDB;
  Ecran.Close;
end ;


procedure TOF_YYADMINBI.OnLoad ;
begin
  Inherited ;
  RemplirCCube;
  RemplirCDatamart;
end ;


procedure TOF_YYADMINBI.OnClose ;
begin
  Inherited ;
  if assigned(TDim) then FreeAndNil(TDim);
  if assigned(TPerso) then FreeAndNil(TPerso);
  if assigned(TRefCubeOlap) then FreeAndNil(TRefCubeOlap);
  if assigned(v_connolap) then FreeAndNil(v_connolap);
  if assigned(myTOB) then FreeAndNil(myTOB);
end ;

procedure TOF_YYADMINBI.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_YYADMINBI.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_YYADMINBI.ChkAnonymousOnClick (Sender:Tobject);
begin
  v_connolap.anonymous := ChkAnonymous.checked;

  SetControlText('ELoginCwas',v_connolap.login);
  SetControlText('EPasswdCwas',v_connolap.password);
  SetControlText('EDomain',v_connolap.domaine);

  ELoginCwas.Enabled:=not(ChkAnonymous.checked);
  EPassWdCwas.Enabled:=not(ChkAnonymous.checked);
  EDomain.Enabled:=not(ChkAnonymous.checked);
end;

procedure TOF_YYADMINBI.ChkAnalDOnClick (Sender:Tobject);
begin
  if ChkanalD.Checked and ChkAnalD.Enabled then
  begin
   InitMoveProgressForm(nil,Ecran.caption, 'Activation de l''analyse décisionnelle, veuillez patienter...', 100, False, true) ;
   MoveCurProgressForm('Activation en cours ...');

   if ActiveOpenOlap then
   begin
     ChkAnalD.Checked:=True;
     ChkAnalD.Enabled:=False;
     v_connolap.activebi := TRUE;
   end
   else
   begin
    PGIBOX(TraduireMemoire('Erreur lors de l''activation de l''analyse décisionnelle. Activation impossible'));
    exit;
   end;
   FiniMoveProgressForm ;
  end;
end;

procedure TOF_YYADMINBI.ECodeCubeOnExit (Sender:Tobject);
var TControle:TOB;
    stWhere:String;
    Qmn:TQuery;
begin
      TControle:=TOB.Create('Controle',nil,-1);
      stwhere := 'WHERE YCO_CODCUBE="'+getControlText('ECodeCube')+'"';
      try
        QMn :=  OpenSQL('SELECT * FROM YCUBEOLAP ' + stWhere,True);
        try
          if not QMn.Eof then
          begin
            PGIBOX('Ce code existe déjà : création impossible');
            SetControlText('ECodeCube','');
            SetFocusControl('ECodeCube');
          end;
        finally
          Ferme(QMn);
        end;
      finally
        TControle.Free;
      end;

end;

procedure TOF_YYADMINBI.CCubeOnExit (Sender:Tobject);
begin
//  ChargeDimensions;
end;

procedure TOF_YYADMINBI.BtnTrtDatamartOnClick (Sender:Tobject);
var   monOlapCell:IOlapCell;
      monFluxXML:WideString;
      monOpenOlapError:IOpenOlapError;
      strerror,strdetail:string;
      ardetail:array of string;
      ndetail:integer;
begin
  SourisSablier;
  monFluxXML:='<?xml version="1.0" ?>'+
        '<dso>'+
          '<server name="'+ v_connolap.ServeurBI +'">'+
            '<database method="update" name="'+ v_connolap.Datamart +'">'+
                  '<process type="processFull"/>'+
            '</database>'+
          '</server>'+
        '</dso>';
  try
  arDetail:=nil;
  monOlapCell:=CoOlapCell.Create;
  InitMoveProgressForm(nil,Ecran.caption, 'Traitement du datamart en cours, veuillez patienter...', 100, False, true) ;
  MoveCurProgressForm('Traitement en cours ...');
  monOlapCell.RunDSO(monFluxXML);
  except
    on E:Exception do
        begin
          monOpenOlapError:=monOlapCell.oxError;
          strError:=monOpenOLAPError.oxDescription +
          ', Level=' + IntToStr(monOpenOLAPError.oxLevel) +
          ', Code=' + IntToStr(monOpenOLAPError.oxCode);

          arDetail:=monOpenOLAPError.oxDetail;
          for nDetail:=VarArrayLowBound(arDetail,1) to VarArrayHighBound(arDetail,1) do
          begin
           strDetail:=arDetail[nDetail];
            if strDetail<>'' then
              strError:=strError + chr(13) + strDetail;
          end;
          PGIBOX(TraduireMemoire('Erreur : ' + strError
                 + monOlapCell.oxrundsooutput
                 ));
          SourisNormale;
          exit;
        end;
    end;
    FiniMoveProgressForm ;
    SourisNormale;
//    PGIBOX(TraduireMemoire('Le datamart '+ v_connolap.Datamart + ' a été entièrement retraité'));
end;

procedure TOF_YYADMINBI.BtnAddPersosOnClick (Sender:Tobject);
var StrError,StrDetail,ChampLibreChoisi,TabletteVal,ChampLibelle,S,Temp:string;
    monfluxxml:widestring;
    monolapcell:iolapcell;
    monopenolaperror:iopenolaperror;
    ardetail:array of string;
    ndetail:integer;
    madimension,maperso:string;
    TCombo:tob;
    qmn:tquery;
    T:TOB;
begin
  if TVDimensions.Selected = Nil then Exit ;
  if TVDimensions.Selected.Level <> 2 then Exit ;


    SourisSablier;
    TCombo:=TOB.Create('TabletteVal',nil,-1);

    S:=TOB(TVDimensions.Selected.Data).GetString('HIE');

    Temp:=ReadTokenSt(S);
    ChampLibreChoisi:=trim(ReadTokenSt(S));

    TabletteVal:=ChampToTT(ChampLibreChoisi);

   Qmn:=OpenSQL('Select DO_CHAMPLIB FROM DECOMBOS WHERE DO_COMBO ="'
      + tabletteval +'"',True) ;
    try
      if not QMn.Eof then TCombo.LoadDetailDB('DO_COMBO','','',QMn,true);
    finally
      Ferme(QMn);
      ChampLibelle := TCombo.detail[0].getvalue('DO_CHAMPLIB');
      TCombo.Free;
    end;

    madimension:=TOB(TVDimensions.Selected.Data).GetString('DIM');

    maperso:=Champtolibelle(ChampLibreChoisi);

    T:=TOB.Create('',TRefCubeOlap,-1);
    T:=TRefCubeOlap.FindFirst(['YRU_YBOGUID','YRU_DIMENSION','YRU_SOURCEKEY']
      ,[v_connolap.guidbase,'['+madimension+']',champlibrechoisi],false);

    if assigned(T) and (T.getstring('YRU_TABLETTEOLAP') <> '')
          then TabletteVal:= T.getString('YRU_TABLETTEOLAP');

    if assigned(T) then
    begin
        T.PutValue('YRU_ATTRIBUT',maperso);
        T.PutValue('YRU_ETAT','X');
        T.AddChampSupValeur('YRU_CHAMPLIBELLE',ChampLibelle);
        T.AddChampSupValeur('YRU_TABLETTE',TabletteVal);
    end;

    ChargeDimensions;

end;

function TOF_YYADMINBI.DsoAdd(dim,table,perso,champlibrechoisi,champlibelle,tabletteval:string):boolean;
var
    StrError,StrDetail:string;
    monfluxxml:widestring;
    monolapcell:iolapcell;
    monopenolaperror:iopenolaperror;
    ardetail:array of string;
    ndetail:integer;
    madimension,maperso:string;
begin
    result:=false;
    if table = '' then table:='dbo_'+dim;
    monFluxXML:= '<?xml version="1.0" ?>'+
    '<dso>'+
      '<server name="'+v_connolap.ServeurBi+'">'+
        '<database method="update" name="'+v_connolap.Datamart+'">'+
          '<dimensions>'+
//            '<dimension method="update" name="['+dim+']">'+
            '<dimension method="update" name="'+dim+'">'+
              '<attributes>'+
                '<attribute method="addifdoesntexist" name="'+perso+'" isaggregatable="true">'+
                  '<columns>'+
                    '<column name="['+table+'].['+ChampLibreChoisi+']" class="key"/>'+
                    '<column name="['+Tabletteval+'].['+ChampLibelle+']" class="name"/>'+
                  '</columns>'+
                '</attribute>'+
              '</attributes>'+
            '</dimension>'+
          '</dimensions>'+
         '</database>'+
      '</server>'+
    '</dso>';

    try
    arDetail:=nil;
    monOlapCell:=CoOlapCell.Create;
    monOpenOlapError:=CoOpenOlapError.Create;

    InitMoveProgressForm(nil,Ecran.caption, 'Création de la personnalisation '+
    perso+' en cours, veuillez patienter...', 100, False, true) ;
    MoveCurProgressForm('Traitement en cours ...');
    monOlapCell.RunDSO(monFluxXML);

    except
    on E:Exception do
        begin
          monOpenOlapError:=monOlapCell.oxError;
          strError:=monOpenOLAPError.oxDescription +
          ', Level=' + IntToStr(monOpenOLAPError.oxLevel) +
          ', Code=' + IntToStr(monOpenOLAPError.oxCode);

          arDetail:=monOpenOLAPError.oxDetail;
          for nDetail:=VarArrayLowBound(arDetail,1) to VarArrayHighBound(arDetail,1) do
          begin
           strDetail:=arDetail[nDetail];
            if strDetail<>'' then
              strError:=strError + chr(13) + strDetail;
          end;
          PGIBOX(TraduireMemoire('Erreur : ' + strError
                 + monOlapCell.oxrundsooutput
                 ));
          SourisNormale;
          exit;
        end;
    end;

    FiniMoveProgressForm ;

    SourisNormale;
    result:=true;

end;

function TOF_YYADMINBI.DsoDel(dim,perso:string):boolean;
var
    StrError,StrDetail:string;
    monfluxxml:widestring;
    monolapcell:iolapcell;
    monopenolaperror:iopenolaperror;
    ardetail:array of string;
    ndetail:integer;
    madimension,maperso:string;
begin
  result:=false;
    monFluxXML:= '<?xml version="1.0" ?>'+
    '<dso>'+
      '<server name="'+v_connolap.ServeurBi+'">'+
        '<database method="update" name="'+v_connolap.Datamart+'">'+
          '<dimensions>'+
            '<dimension method="update" name="'+dim+'">'+
              '<attributes>'+
                '<attribute method="remove" name="'+perso+'"/>'+
              '</attributes>'+
            '</dimension>'+
          '</dimensions>'+
         '</database>'+
      '</server>'+
    '</dso>';

    try
    arDetail:=nil;
    monOlapCell:=CoOlapCell.Create;
    monOpenOlapError:=CoOpenOlapError.Create;

    InitMoveProgressForm(nil,Ecran.caption, 'Suppression de la personnalisation '+
    perso+' en cours, veuillez patienter...', 100, False, true) ;
    MoveCurProgressForm('Traitement en cours ...');
    monOlapCell.RunDSO(monFluxXML);

    except
    on E:Exception do
        begin
          monOpenOlapError:=monOlapCell.oxError;
          strError:=monOpenOLAPError.oxDescription +
          ', Level=' + IntToStr(monOpenOLAPError.oxLevel) +
          ', Code=' + IntToStr(monOpenOLAPError.oxCode);

          arDetail:=monOpenOLAPError.oxDetail;
          for nDetail:=VarArrayLowBound(arDetail,1) to VarArrayHighBound(arDetail,1) do
          begin
           strDetail:=arDetail[nDetail];
            if strDetail<>'' then
              strError:=strError + chr(13) + strDetail;
          end;
          PGIBOX(TraduireMemoire('Erreur : ' + strError
                 + monOlapCell.oxrundsooutput
                 ));
          SourisNormale;
          exit;
        end;
    end;

    FiniMoveProgressForm ;

//    PGIBOX(TraduireMemoire('La personnalisation a bien été supprimée'));

{    ExecuteSQL(
      'UPDATE YREFCUBEOLAP SET YRU_ETAT = "" WHERE YRU_DIMENSION ="[' +madimension+']" AND YRU_ATTRIBUT="'+maperso+'"'
      );}

  result:=true;
end;

procedure TOF_YYADMINBI.BtnDelPersosOnClick (Sender:Tobject);
var
  madimension,maperso,strerror,strdetail,S,Temp,ChampLibreChoisi:string;
  monfluxxml:widestring;
  monolapcell:iolapcell;
  monopenolaperror:iopenolaperror;
  ardetail:array of string;
  ndetail:integer;
  T:TOB;
begin
  if TVPersos.Selected = Nil then Exit ;
  if TVPersos.Selected.Level <> 2 then Exit ;
  S:=TOB(TVPersos.Selected.Data).GetString('HIE');

  Temp:=ReadTokenSt(S);
  ChampLibreChoisi:=trim(ReadTokenSt(S));
  maperso:=Champtolibelle(ChampLibreChoisi);

  madimension:=TOB(TVPersos.Selected.Data).GetString('DIM');

  T:=TOB.Create('',TRefCubeOlap,-1);
  T:=TRefCubeOlap.FindFirst(['YRU_YBOGUID','YRU_DIMENSION','YRU_SOURCEKEY']
    ,[v_connolap.guidbase,'['+madimension+']',champlibrechoisi],false);

  if assigned(T) then
  begin
      T.PutValue('YRU_ATTRIBUT',maperso);
      T.PutValue('YRU_ETAT','');
  end;

  ChargeDimensions;

end;

procedure TOF_YYADMINBI.BtnTestConnectOnClick (Sender:Tobject);
var stwhere:string;
    Qmn:TQuery;
begin
  if GetControlText('EServeurbi') <> '' then
  begin
    v_connolap.codecube:=GetControlText('ECodeCube');
    v_connolap.serveurbi:=GetControlText('EServeurBi');
    v_connolap.datamart:=CDatamart.Text;
    v_connolap.cube:=CCube.Text;
    v_connolap.libcube:=GetControlText('ELibCube');
    v_connolap.serveurcwas:=GetControlText('EServeurCwas');
    v_connolap.login:=GetControlText('ELoginCwas');
    v_connolap.password:=GetControlText('EPasswdCwas');
    v_connolap.domaine:=GetControlText('EDomain');

    myTOB.Detail[0].InsertOrUpdateDb;

    if (v_connolap.ServeurBi <> '') and (v_connolap.CodeCube <> '') then
    begin
      TRefCubeOlap:=TOB.Create('Referentiel',nil,-1);
      TRefCubeOlap.AddChampSup('YRU_CHAMPLIBELLE',true);
      TRefCubeOlap.AddChampSup('YRU_TABLETTE',true);
      stwhere := 'WHERE YRU_YBOGUID="'+v_connolap.guidbase+'"';
      try
        QMn :=  OpenSQL('SELECT * FROM YREFCUBEOLAP ' + stWhere,True);
        try
          if not QMn.Eof then TRefCubeOlap.LoadDetailDB('YREFCUBEOLAP','','',QMn,true);
        finally
          Ferme(QMn);
        end;
      finally
      end;
      ChargeDimensions;
      PGIBOX (TraduireMemoire('Test de connexion réalisé avec succès'));
    end;
  end;
end;

procedure TOF_YYADMINBI.ChargeDimensions;
var monOlapCell,monOlapCelldim:IOlapCell;
    iInd,iIndrefcube:integer;
    ardimensions:olevariant;
    st,strerror,stdim:string;
    monOpenOLAPErrordim:iopenolaperror;
    TDimDetail,Thie:TOB;
    TPersoDetail,ThiePerso:TOB;
begin
  ardimensions:=okConnectCube;
  if not VarIsArray(arDimensions) then exit;

// on remplit la liste des dimensions

  if assigned(TRefCubeOlap) and (TRefCubeOlap.FillesCount(100)=0) then exit;
  if (GetControlText('EServeurBI') = '') or (CDatamart.Text = '') or (CCube.Text = '') then exit;

  if assigned(TRefCubeOlap) then
  begin
    st:='';
    TVDimensions.CleanupInstance;
    TVPersos.CleanupInstance;
    if assigned(TDim) then TDim.ClearDetail
      else
      begin
        TDim:=TOB.Create('EnteteDim',nil,-1);
        TDim.AddChampSupValeur('DIM','Dimension',false);
        TDim.AddChampSupValeur('HIE','Hierarchie',false);
      end;

    if assigned(TPerso) then TPerso.ClearDetail
      else
      begin
        TPerso:=TOB.Create('EntetePerso',nil,-1);
        TPerso.AddChampSupValeur('DIM','Dimension',false);
        TPerso.AddChampSupValeur('HIE','Hierarchie',false);
      end;

    for iInd:= vararraylowbound(ardimensions,1) to vararrayhighbound(ardimensions,1) do
    begin
  //    LstDimensions.Items.Add(ardimensions[iInd]);
      stdim := stringreplace(stringreplace(ardimensions[iInd],'[','',[rfReplaceAll]),']','',[rfReplaceAll]);
      TDimDetail:=TOB.Create('DetailDim',TDim,-1);
      TDim.Detail[iInd].AddChampSupValeur('DIM',stdim);

      TPersoDetail:=TOB.Create('DetailPerso',TPerso,-1);
      TPerso.Detail[iInd].AddChampSupValeur('DIM',stdim);

      MoveCurProgressForm('Dimensions : '+ ardimensions[iInd]);

      for iIndrefcube:=0 to TRefCubeOlap.detail.count - 1 do
      begin
           if (TRefCubeOlap.detail[iIndrefcube].getvalue('YRU_DIMENSION') = ardimensions[iInd]) then
           begin
             Thie:=TOB.Create('Hierarchie',TDimDetail,-1);
             THiePerso:=TOB.Create('Hierarchie',TPersoDetail,-1);
               if TRefCubeOlap.detail[iIndrefcube].getvalue('YRU_ETAT') <> 'X'
                  then
                  begin
                   TDimDetail.Detail[iIndrefcube].AddChampSupValeur('DIM',stdim);
                   TDimDetail.Detail[iIndrefcube].AddChampSupValeur('HIE',ChampToLibelle(TRefCubeOlap.detail[iIndrefcube].getvalue('YRU_SOURCEKEY')) +
                   ' ; ' + TRefCubeOlap.detail[iIndrefcube].getvalue('YRU_SOURCEKEY'));
                  end
                  else
                  begin
                   TPersoDetail.Detail[iIndrefcube].AddChampSupValeur('DIM',stdim);
                   TPersoDetail.Detail[iIndrefcube].AddChampSupValeur('HIE',ChampToLibelle(TRefCubeOlap.detail[iIndrefcube].getvalue('YRU_SOURCEKEY')) +
                   ' ; ' + TRefCubeOlap.detail[iIndrefcube].getvalue('YRU_SOURCEKEY'));
                  end;
           end;
        end;
    end;

    TDim.PutTreeView(TVDimensions,nil,'"Dimension :";DIM;HIE',-1,-1);
    TPerso.PutTreeView(TVPersos,nil,'"Dimension :";DIM;HIE',-1,-1);
    TVDimensions.FullExpand;
    TVPersos.FullExpand;
  end;
//  if assigned(TDim) then TDim.Free;
//  if assigned(TPerso) then TPerso.Free;
end;

function TOF_YYADMINBI.ActiveOpenOlap:boolean;
var FileName:string;
begin
  FileName:=ExtractFilePath(Application.Exename)+ 'OpenOlapInstallFull.exe';
  if FileExists(Filename) then
  begin
    FileExecAndWait(FileName + ' /Q');
//    DeleteFile(FileName);
    result:=True;
  end
  else result :=false;
end;

procedure TOF_YYADMINBI.CDatamartOnClick(Sender:TObject);
begin
  RemplirCCube;
end;

procedure TOF_YYADMINBI.EServeurBiOnExit(Sender:TObject);
begin
  RemplirCDatamart;
end;

procedure TOF_YYADMINBI.CDatamartOnExit(Sender:TObject);
begin
  RemplirCCube;
end;

procedure TOF_YYADMINBI.RemplirCDatamart;
var iInd:integer;
    ardatamarts:olevariant;
    st,temp,tempinlist:string;
begin
//  if CDatamart.Text <> '' then exit;
  if getcontroltext('Eserveurbi') = '' then exit;
  ardatamarts:=okConnectSrv;
  if not VarIsArray(ardatamarts) then exit;

  InitMoveProgressForm(nil,Ecran.caption, 'Chargement des datamarts disponibles, veuillez patienter...', 100, False, true) ;
  MoveCurProgressForm('Connexion en cours ...');
// on remplit la liste des datamarts
//  monOlapCell:=CoOlapCell.create;
{  monOlapCell.Connect('',
    '','','',
    getControlText('EServeurbi'),'',false);}

  MoveCurProgressForm('Connexion effectuée ...');
//  ardatamarts:=monOlapCell.oxdatamarts;
  st:='';
  temp:=CDatamart.Text;
  tempinlist:='';
  CDatamart.clear;
  for iInd:= vararraylowbound(ardatamarts,1) to vararrayhighbound(ardatamarts,1) do
  begin
    CDatamart.Items.Add(ardatamarts[iInd]);
    if temp=ardatamarts[iInd] then tempinlist:=temp;
    MoveCurProgressForm('Datamart : '+ ardatamarts[iInd]);
  end;
  if tempinlist <> '' then CDatamart.Text:=tempinlist;

  FiniMoveProgressForm ;
end;

procedure TOF_YYADMINBI.RemplirCCube;
var monOlapCell:IOlapCell;
    iInd:integer;
    arcubes:olevariant;
    st,temp,tempinlist:string;
begin
//    if CCube.Text <> '' then exit;
    if getControlText('Eserveurbi') = '' then exit;
    if CDatamart.Text = '' then exit;
    arcubes:=okConnectDatamart;
    if not VarIsArray(arcubes)then exit;
  //  if (Action='ACTION=MODIFICATION') and (v_connolap.Datamart <> '') then exit;

    InitMoveProgressForm(nil,Ecran.caption, 'Chargement des cubes disponibles, veuillez patienter...', 100, False, true) ;
    MoveCurProgressForm('Connexion en cours ...');
  // on remplit la liste des cubes
//    monOlapCell:=CoOlapCell.create;
    {monOlapCell.Connect('',
      '','','',
      getcontroltext('Eserveurbi'),CDatamart.Text,false);}

    MoveCurProgressForm('Connexion effectuée ...');
//    arcubes:=monOlapCell.oxcubes;
    st:='';
    temp:=CCube.Text;
    tempinlist:='';
    CCube.clear;
    for iInd:= vararraylowbound(arcubes,1) to vararrayhighbound(arcubes,1) do
    begin
      CCube.Items.Add(arcubes[iInd]);
      if temp=arcubes[iInd] then tempinlist:=temp;
      MoveCurProgressForm('Cube : '+ arcubes[iInd]);
    end;
    if tempinlist<> '' then CCube.Text:=tempinlist;

    FiniMoveProgressForm ;
end;

function TOF_YYADMINBI.okConnectDatamart : olevariant;
var monOlapCell:IOlapCell;
begin
  result:='';
  monOlapCell:=CoOlapCell.create;
  try
   monOlapCell.Connect(GetControlText('EServeurCwas'),
      GetControlText('ELoginCwas'),GetControlText('EPasswdCwas'),GetControlText('EDomain'),
      GetControlText('EServeurBi'),CDatamart.Text,false);
    except
    on E:Exception do
      begin
        result:='';
        PGIBOX(TraduireMemoire('Echec de connexion à la base OLAP : veuillez vérifier votre paramétrage !'));
        exit;
      end;
  end;
  if monOlapCell.oxIsConnected then result:=monOlapCell.oxCubes;
  CDatamart.Enabled:=True;
  CCube.Enabled:=True;
end;


function TOF_YYADMINBI.OkConnectSrv : olevariant;
var monOlapCell:IOlapCell;
begin
  result:='';
  monOlapCell:=CoOlapCell.create;
  try
    monOlapCell.Connect(GetControlText('EServeurCwas'),
      GetControlText('ELoginCwas'),GetControlText('EPasswdCwas'),GetControlText('EDomain'),
      GetControlText('EServeurBi'),'',false);
    except
    on E:Exception do
      begin
        result:='';
        PGIBOX(TraduireMemoire('Echec de connexion au serveur : veuillez vérifier votre paramétrage !'));
        exit;
      end;
  end;
  if monOlapCell.oxIsConnected then result:=monOlapCell.oxDatamarts;
  CDatamart.Enabled:=True;
end;

function TOF_YYADMINBI.OkConnectCube : olevariant;
var monOlapCell:IOlapCell;
begin                                                    
  result:='';
  monOlapCell:=CoOlapCell.create;
  try
    monOlapCell.Connect(GetControlText('EServeurCwas'),
      GetControlText('ELoginCwas'),GetControlText('EPasswdCwas'),GetControlText('EDomain'),
      GetControlText('EServeurBi'),CDatamart.Text,false);
    monOlapCell.oxCube:=CCube.Text;
    except
    on E:Exception do
      begin
        result:='';
        PGIBOX(TraduireMemoire('Echec de connexion au cube : veuillez vérifier votre paramétrage !'));
        exit;
      end;
  end;
  if monOlapCell.oxIsConnected then result:=monOlapCell.oxDimensions;
  CDatamart.Enabled:=True;
end;


procedure monobjet.SetCube(const Value: string);
begin
   if (myTOB.Detail.Count > 0) then myTOB.Detail[0].PutValue('YCO_CUBE',value);
end;

procedure monobjet.SetDataMart(const Value: string);
begin
   if (myTOB.Detail.Count > 0) then myTOB.Detail[0].PutValue('YCO_DATAMART',value);
end;

procedure monobjet.SetDomaine(const Value: string);
begin
   if (myTOB.Detail.Count > 0) then myTOB.Detail[0].PutValue('YCO_DOMCWAS',value);
end;

procedure monobjet.SetLogin(const Value: string);
begin
   if (myTOB.Detail.Count > 0) then myTOB.Detail[0].PutValue('YCO_LOGINCWAS',value);
end;

procedure monobjet.SetPassword(const Value: string);
begin
   if (myTOB.Detail.Count > 0) then myTOB.Detail[0].PutValue('YCO_PWDCWAS',value);
end;

procedure monobjet.SetactiveBi(const Value: boolean);
begin
  Setparamsoc('SO_ACTIVEBI',Value);
end;

procedure monobjet.Setanonymous(const Value: boolean);
begin
  Login:='anonymous';
  Password:='';
  Domaine:='';
end;

procedure monobjet.SetServeurBi(const Value: string);
begin
   if (myTOB.Detail.Count > 0) then myTOB.Detail[0].PutValue('YCO_SERVEURBI',value);
end;

procedure monobjet.SetServeurCwas(const Value: string);
begin
   if (myTOB.Detail.Count > 0) then myTOB.Detail[0].PutValue('YCO_SERVEURCWAS',value);
end;

procedure monobjet.SetCodeCube(const Value: string);
begin
  FcodeCube := Value;
   if (myTOB.Detail.Count > 0) then myTOB.Detail[0].PutValue('YCO_CODCUBE',FCodeCube);
end;

procedure monobjet.SetLibCube(const Value: string);
begin
   if (myTOB.Detail.Count > 0) then myTOB.Detail[0].PutValue('YCO_LIBCUBE',Value);
end;

procedure monobjet.SetGuidBase(const Value: string);
begin
   if (myTOB.Detail.Count > 0) then myTOB.Detail[0].PutValue('YCO_YBOGUID',Value);
end;

function monobjet.GetLibCube: string;
begin
  if (myTOB.Detail.Count > 0) and myTOB.Detail[0].ExistDB
    then result := myTOB.Detail[0].getString('YCO_LIBCUBE');
end;

function monobjet.GetCodeCube: string;
begin
  result := FCodeCube;
end;

function monobjet.GetLogin: string;
begin
  if (myTOB.Detail.Count > 0) and (myTOB.Detail[0].getString('YCO_LOGINCWAS') = '') then
    result := 'anonymous'
  else
    result := myTOB.Detail[0].getString('YCO_LOGINCWAS');
end;

function monobjet.GetServeurBI: string;
begin
  if (myTOB.Detail.Count > 0) and myTOB.Detail[0].ExistDB
    then result := myTOB.Detail[0].getString('YCO_SERVEURBI');
end;

function monobjet.GetActiveBI: boolean;
begin
    result := getparamsocsecur('SO_ACTIVEBI','');
end;

function monobjet.Getanonymous: boolean;
begin
  if (myTOB.Detail.Count > 0) and (myTOB.Detail[0].getString('YCO_LOGINCWAS') = 'anonymous')
    then result := true
    else result := false;
end;

function monobjet.GetDatamart: string;
begin
  if (myTOB.Detail.Count > 0) and myTOB.Detail[0].ExistDB
    then result := myTOB.Detail[0].getString('YCO_DATAMART');
end;

function monobjet.GetCube: string;
begin
  if (myTOB.Detail.Count > 0) and myTOB.Detail[0].ExistDB
    then result := myTOB.Detail[0].getString('YCO_CUBE');
end;

function monobjet.GetServeurCwas: string;
begin
  if (myTOB.Detail.Count > 0) and myTOB.Detail[0].ExistDB
    then result := myTOB.Detail[0].getString('YCO_SERVEURCWAS');
end;

function monobjet.GetPassword: string;
begin
  if (myTOB.Detail.Count > 0) and myTOB.Detail[0].ExistDB
    then result := myTOB.Detail[0].getString('YCO_PWDCWAS');
end;

function monobjet.GetDomaine: string;
begin
  if (myTOB.Detail.Count > 0) and myTOB.Detail[0].ExistDB
    then result := myTOB.Detail[0].getString('YCO_DOMCWAS');
end;

function monobjet.GetGuidBase: string;
begin
  if (myTOB.Detail.Count > 0) and myTOB.Detail[0].ExistDB
    then result := myTOB.Detail[0].getString('YCO_YBOGUID');
end;

constructor monobjet.Create(AOwner: TComponent);
var iCount:integer;
begin
  FOwner := AOwner;
  iCount := 0;
  if (myTOB.Detail.Count > 0) and myTOB.Detail[0].ExistDB
    then
  begin
    FCodeCube := myTOB.Detail[0].getString('YCO_CODCUBE');
  end;
end;

destructor monobjet.Destroy;
begin
  inherited;
end;


Initialization
  registerclasses ( [ TOF_YYADMINBI ] ) ;
end.
