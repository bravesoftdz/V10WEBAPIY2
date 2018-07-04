{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 05/01/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : YYADMUSRBI ()
Mots clefs ... : TOF;YYADMUSRBI
*****************************************************************}
Unit YYADMUSRBI_TOF ;

Interface

Uses
{$IFDEF VER150}
  variants,
{$ENDIF}
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
     uTob, 
     forms,
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF,
     Hsysmenu,
     HTB97,
     OPENOLAP4Lib_TLB;

Type
  TOF_YYADMUSRBI = Class (TOF)
    procedure OnNew                    ; override ;

    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    DroitsTOB,FiltresTOB,YCOTOB,TOBUSR:TOB;
    DroitRow,FiltreRow:TOB;
    GDroits,GFiltres:THGRID;
    CCube,CDimension,CHierarchy,CMember:THValComboBox;
    CBAnonymous:TCheckBox;
    ELogin,EPassWord,EDomain,EGuidDroits,EGuidBase,EValue,ECodeUsr,
      EUsr,EGroupUsr,EFonctUsr:THEdit;
    BValidDroit,BInsertDroit,BDeleteDroit,BDefaitDroit,
      BValidFiltre,BInsertFiltre,BDeleteFiltre,BDefaitFiltre,
      BDroitGrp,BFiltreGrp:TToolbarButton97;
    HMTrad: THSystemMenu;
    procedure GDroitsOnClick(Sender:TObject);
    procedure GDroitsOnDblClick(Sender:TObject);
    procedure CBAnonymousOnClick(Sender:TObject);    
    procedure GFiltresOnClick(Sender:TObject);
    procedure GFiltresOnDblClick(Sender:TObject);
    procedure BValidDroitOnClick(Sender:TObject);
    procedure BInsertDroitOnClick(Sender:TObject);
    procedure BDeleteDroitOnClick(Sender:TObject);
    procedure BDefaitDroitOnClick(Sender:TObject);
    procedure BValidFiltreOnClick(Sender:TObject);
    procedure BInsertFiltreOnClick(Sender:TObject);
    procedure BDeleteFiltreOnClick(Sender:TObject);
    procedure BDefaitFiltreOnClick(Sender:TObject);
    procedure BDroitGrpOnClick(Sender:TObject);
    procedure BFiltreGrpOnClick(Sender:TObject);
    procedure EnableCtrlsDroits;
    procedure DisableCtrlsDroits;
    procedure VideCtrlsDroits;
    procedure EnableCtrlsFiltres;
    procedure DisableCtrlsFiltres;
    procedure VideCtrlsFiltres;
    procedure ChargeCDimension(CodeCube:string);
    procedure ChargeCHierarchy(Dimension:string);
    procedure ChargeCMember(Hierarchy:string);
    procedure CDimensionOnClick(Sender:TObject);
    procedure CHierarchyOnClick(Sender:TObject);
    function FormatteMember(Member:string) : string;
    procedure AlimGDroits(ArgCodeUsr:String);
    procedure AlimGFiltres(ArgCodeUsr,ArgCodeCube:String);
end ;

Implementation

uses  utobdebug;

procedure TOF_YYADMUSRBI.DisableCtrlsDroits ;
begin
  CCube.Enabled:=False;
  CBAnonymous.Enabled:=False;
  ELogin.Enabled:=False;
  EPassWord.Enabled:=False;
  EDomain.Enabled:=False;
end;

procedure TOF_YYADMUSRBI.EnableCtrlsDroits ;
begin
  CCube.Enabled:=True;
  CBAnonymous.Enabled:=True;
  ELogin.Enabled := not(CBAnonymous.Checked);
  EPassWord.Enabled := not(CBAnonymous.Checked);
  EDomain.Enabled := not(CBAnonymous.Checked);
end;

procedure TOF_YYADMUSRBI.VideCtrlsDroits ;
begin
  SetControlText('EGUIDDROITS','');
  SetControlText('EGUIDBASE','');  
  CCube.Text:='';
  SetControlText('ELOGIN','');
  SetControlText('EPASSWORD','');
  SetControlText('EDOMAIN','');
  CbAnonymous.Checked := False;
end;

procedure TOF_YYADMUSRBI.DisableCtrlsFiltres ;
begin
  CDimension.Enabled:=False;
  CHierarchy.Enabled:=False;
  CMember.Enabled:=False;
end;

procedure TOF_YYADMUSRBI.VideCtrlsFiltres ;
begin
  CDimension.Text:='';
  CHierarchy.Text:='';
  CMember.Text:='';
end;

procedure TOF_YYADMUSRBI.EnableCtrlsFiltres ;
begin
  CDimension.Enabled:=True;
  CHierarchy.Enabled:=True;
  CMember.Enabled:=True;
end;

procedure TOF_YYADMUSRBI.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_YYADMUSRBI.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_YYADMUSRBI.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_YYADMUSRBI.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_YYADMUSRBI.OnArgument (S : String ) ;
Var ArgCodeUsr,StWhere:string;
    Qmn : TQuery;
begin
  Inherited ;
  ArgCodeUsr:=S;
  if not assigned(getControl('GDROITS')) then exit;

  GDroits:=THGrid(GetControl('GDROITS'));
  GDroits.OnClick:=GDroitsOnClick;
  GDroits.OnDblClick:=GDroitsOnDblClick;

  GFiltres:=THGrid(GetControl('GFILTRES'));
  GFiltres.OnClick:=GFiltresOnClick;
  GFiltres.OnDblClick:=GFiltresOnDblClick;

  CCube:=THValComboBox(GetControl('CCUBE'));
  CDimension:=THValComboBox(GetControl('CDIMENSION'));
  CDimension.OnClick:=CDimensionOnClick;
  CHierarchy:=THValComboBox(GetControl('CHIERARCHY'));
  CHierarchy.OnClick:=CHierarchyOnClick;
  CMember:=THValComboBox(GetControl('CMEMBER'));

  BValidDroit:=TToolbarButton97(getControl('BValidDroit'));
  BInsertDroit:=TToolbarButton97(getControl('BInsertDroit'));
  BDeleteDroit:=TToolbarButton97(getControl('BDeleteDroit'));
  BDefaitDroit:=TToolbarButton97(getControl('BDefaitDroit'));
  BValidFiltre:=TToolbarButton97(getControl('BValidFiltre'));
  BInsertFiltre:=TToolbarButton97(getControl('BInsertFiltre'));
  BDeleteFiltre:=TToolbarButton97(getControl('BDeleteFiltre'));
  BDefaitFiltre:=TToolbarButton97(getControl('BDefaitFiltre'));
  BDroitGrp:=TToolbarButton97(getControl('BDroitGrp'));
  BFiltreGrp:=TToolbarButton97(getControl('BFiltreGrp'));
  BValidDroit.OnClick:=BValidDroitOnClick;
  BInsertDroit.OnClick:=BInsertDroitOnClick;
  BDeleteDroit.OnClick:=BDeleteDroitOnClick;
  BDefaitDroit.OnClick:=BDefaitDroitOnClick;
  BValidFiltre.OnClick:=BValidFiltreOnClick;
  BInsertFiltre.OnClick:=BInsertFiltreOnClick;
  BDeleteFiltre.OnClick:=BDeleteFiltreOnClick;
  BDefaitFiltre.OnClick:=BDefaitFiltreOnClick;
  BDroitGrp.OnClick:=BDroitGrpOnClick;
  BFiltreGrp.OnClick:=BFiltreGrpOnClick;

  CBAnonymous:=TCheckBox(getControl('CBANONYMOUS'));
  CBAnonymous.OnClick:=CBAnonymousOnClick;
  ELogin:=THEdit(getControl('ELOGIN'));
  EPassWord:=THEdit(getControl('EPASSWORD'));
  EDomain:=THEdit(getControl('EDOMAIN'));
  EGuidDroits:=THEdit(getControl('EGUIDDROITS'));
  EGuidBase:=THEdit(getControl('EGUIDBASE'));  
  EValue:=THEdit(getControl('EVALUE'));
  ECodeUsr:=THEdit(getControl('ECODEUSR'));
  EUsr:=THEdit(getControl('EUSR'));
  EGroupUsr:=THEdit(getControl('EGROUPUSR'));
  EFonctUsr:=THEdit(getControl('EFONCTUSR'));

  TOBUSR := TOB.Create('MonUser',nil,-1);
  stwhere := 'LEFT JOIN USERGRP ON US_GROUPE=UG_GROUPE WHERE US_UTILISATEUR="'+ArgCodeUsr+'"';
  QMn :=  OpenSQL('SELECT * FROM UTILISAT ' + stWhere,True);
  try
    if not QMn.Eof then
        TOBUSR.LoadDetailDB('MonUser','','',QMn,true);
  finally
    Ferme(QMn);
  end;

  if (TOBUSR.Detail.Count > 0) then
  begin
    SetControlText('ECODEUSR',TOBUSR.Detail[0].GetString('US_UTILISATEUR'));
    SetControlText('EUSR',TOBUSR.Detail[0].GetString('US_LIBELLE'));
    SetControlText('EGROUPUSR',TOBUSR.Detail[0].GetString('UG_LIBELLE'));
    SetControlText('EFONCTUSR',TOBUSR.Detail[0].GetString('US_FONCTION'));
  end;

  AlimGdroits(ArgCodeUsr);

  if not assigned(getControl('GFILTRES')) then exit;

  CCube.Enabled:=False;
  DisableCtrlsDroits;
  DisableCtrlsFiltres;

end ;

procedure TOF_YYADMUSRBI.AlimGDroits(ArgCodeUsr:String);
var stWhere : string;
    QMn:TQuery;
begin
  DroitsTOB:=TOB.Create('YUSRCUBEOLAP',nil,-1);
  if ArgCodeUsr <> ''
    then
    begin
      stwhere := 'WHERE YUC_USER="'+ArgCodeUsr+'" ORDER BY YUC_CODCUBE';
      QMn :=  OpenSQL('SELECT * FROM YUSRCUBEOLAP LEFT JOIN YCUBEOLAP ON YUC_CODCUBE = YCO_CODCUBE ' + stWhere,True);
      try
        if not QMn.Eof then
            DroitsTOB.LoadDetailDB('YUSRCUBEOLAP','','',QMn,true);
      finally
        Ferme(QMn);
      end;
      DroitsTOB.PutGridDetail(GDroits,True,True,'YUC_CODCUBE;YCO_CUBE;YUC_ANONYMOUS;YUC_LOGINCWAS;YUC_DOMCWAS',True);
      HMTrad.ResizeGridColumns(GDroits);
    end;
  GDroitsOnClick(nil);
end;

procedure TOF_YYADMUSRBI.OnClose ;
begin
  Inherited ;
  if assigned(TOBUSR) then  FreeAndNil(TOBUSR);
  if assigned(DroitRow) then  FreeAndNil(DroitRow);
  if assigned(FiltreRow) then  FreeAndNil(FiltreRow);
  if assigned(DroitsTOB) then  FreeAndNil(DroitsTOB);
  if assigned(FiltresTOB) then  FreeAndNil(FiltresTOB);
  if assigned(YCOTOB) then  FreeAndNil(YCOTOB);
end ;

procedure TOF_YYADMUSRBI.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_YYADMUSRBI.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_YYADMUSRBI.GDroitsOnClick (Sender : TObject) ;
var stWhere:string;
    Qmn:TQuery;
    TOBBaseOlap : TOB;
begin
  if not assigned(FiltresTOB) then FiltresTOB:=TOB.Create('YFILBASEOLAP',nil,-1);
  DroitRow := TOB(GDroits.Objects[0,GDroits.Row]);
  if DroitRow=nil then exit;
  SetControlText('EGUIDDROITS',DroitRow.GetString('YUC_GUID'));

  TOBBaseOlap:=TOB.Create('YBASEOLAP',nil,-1);
  stwhere := 'LEFT JOIN YCUBEOLAP ON YCO_YBOGUID=YBO_GUID '+
    'WHERE YCO_CODCUBE="'+DroitRow.getString('YCO_CODCUBE')+'"';
  QMn :=  OpenSQL('SELECT * FROM YBASEOLAP ' + stWhere,True);
  try
    if not QMn.Eof then
        TOBBaseOlap.LoadDetailDB('MaBase','','',QMn,true);
  finally
    Ferme(QMn);
  end;
  if (TOBBaseOlap.Detail.Count > 0) then
    SetControlText('EGUIDBASE',TOBBaseOlap.Detail[0].GetString('YBO_GUID'));

  CCube.Text:=DroitRow.getString('YUC_CODCUBE');
  SetControlText('ELOGIN',DroitRow.GetString('YUC_LOGINCWAS'));
  SetControlText('EPASSWORD',DroitRow.GetString('YUC_PWDCWAS'));
  SetControlText('EDOMAIN',DroitRow.GetString('YUC_DOMCWAS'));
  if DroitRow.GetBoolean('YUC_ANONYMOUS')
    then CbAnonymous.Checked := True
    else CbAnonymous.Checked := False;

  AlimGFiltres(DroitRow.getString('YUC_USER'),DroitRow.getString('YUC_CODCUBE'));
end ;

procedure TOF_YYADMUSRBI.AlimGFiltres(ArgCodeUsr,ArgCodeCube:String);
var stWhere:string;
    Qmn:TQuery;
    TOBCubeOlap:TOB;
    FGuidBase:string;
begin
  FiltresTOB.ClearDetail;

//js1 on va chercher la guid de la base contenant le cube courant
  TOBCubeOlap:=TOB.Create('YCUBEOLAP',nil,-1);
  try
    TOBCubeOlap.LoadDetailDB('YCUBEOLAP','"'+ArgCodeCube+'"','',nil,False);
    try
      if (TOBCubeOlap.Detail.Count > 0) and TOBCubeOlap.Detail[0].ExistDB then
        FGuidBase:=TOBCubeOlap.Detail[0].getString('YCO_YBOGUID');
    except
      PGIBOX(TraduireMemoire('Erreur dans le paramétrage du cube !'));
      Abort;
    end;
  finally
    TOBCubeOlap.Free;
  end;
  
  stwhere := 'WHERE YFL_USER="'+ArgCodeUsr+'" AND '
    + 'YFL_YBOGUID = "' + FGuidBase +'"';
  QMn := OpenSQL('SELECT * FROM YFILBASEOLAP ' + stWhere,True);
  try
    if not QMn.Eof then
        FiltresTOB.LoadDetailDB('YFILBASEOLAP','','',QMn,true);
  finally
    Ferme(QMn);
  end;
  FiltresTOB.PutGridDetail(GFiltres,True,True,'YFL_HIERARCHY;YFL_VALUE',True);
  HMTrad.ResizeGridColumns(GFiltres);

  GFiltresOnClick(nil);
end;

procedure TOF_YYADMUSRBI.ChargeCDimension(CodeCube:string);
var monOlapcell:iOlapcell;
    iInd:integer;
    ardetail:oleVariant;
begin
  YCOTOB:=TOB.Create('YCUBEOLAP',nil,-1);
  YCOTOB.LoadDetailDB('YCUBEOLAP','"'+CodeCube+'"','',nil,False);
  if YCOTOB.Detail.Count = 0 then exit;

  CDimension.Clear;
  monOlapCell:=CoOlapCell.Create;
  try
    with YCOTOB.Detail[0] do
    begin
      monOlapCell.Connect(GetString('YCO_SERVEURCWAS'),
              GetString('YCO_LOGINCWAS'),
              GetString('YCO_PWDCWAS'),
              GetString('YCO_DOMCWAS'),
              GetString('YCO_SERVEURBI'),
              GetString('YCO_DATAMART'),false);
      monOlapCell.oxCube:=GetString('YCO_CUBE');
    end;
  except
  on E : Exception do
    begin
        PGIBOX(TraduireMemoire('Echec de connexion au cube : veuillez vérifier votre paramétrage !'));
        exit;
    end;
  end;

  if not monOlapCell.oxIsConnected then exit;
  arDetail:=monOlapCell.Get_oxDimensions;
  begin
    CDimension.Clear;
    for iInd := vararraylowbound(ardetail,1) to vararrayhighbound(ardetail,1) do
    begin
      CDimension.Items.Add(ardetail[iInd]);
    end;
  end;
end;

procedure TOF_YYADMUSRBI.ChargeCHierarchy(Dimension:string);
var monOlapcell:iOlapcell;
    iInd:integer;
    ardetail:oleVariant;
begin
  CHierarchy.Clear;
  monOlapCell:=CoOlapCell.Create;
  if YCOTOB.Detail.Count = 0 then exit;
  try
    with YCOTOB.Detail[0] do
    begin
      monOlapCell.Connect(GetString('YCO_SERVEURCWAS'),
            GetString('YCO_LOGINCWAS'),
            GetString('YCO_PWDCWAS'),
            GetString('YCO_DOMCWAS'),
            GetString('YCO_SERVEURBI'),
            GetString('YCO_DATAMART'),false);
      monOlapCell.oxCube:=GetString('YCO_CUBE');
    end;
  except
  on E : Exception do
    begin
        PGIBOX(TraduireMemoire('Echec de connexion au cube : veuillez vérifier votre paramétrage !'));
        exit;
    end;
  end;

  if not monOlapCell.oxIsConnected then exit;
  arDetail:=monOlapCell.Get_oxHierarchies(Dimension);
  begin
    CHierarchy.Clear;
    for iInd := vararraylowbound(ardetail,1) to vararrayhighbound(ardetail,1) do
    begin
      CHierarchy.Items.Add(ardetail[iInd]);
    end;
  end;
end;

procedure TOF_YYADMUSRBI.ChargeCMember(Hierarchy:string);
var monOlapcell:iOlapcell;
    iInd:integer;
    ardetail:oleVariant;
begin
  CMember.Clear;
  monOlapCell:=CoOlapCell.Create;
  if YCOTOB.Detail.Count = 0 then exit;

  try
  with YCOTOB.Detail[0] do
  begin
    monOlapCell.Connect(GetString('YCO_SERVEURCWAS'),
            GetString('YCO_LOGINCWAS'),
            GetString('YCO_PWDCWAS'),
            GetString('YCO_DOMCWAS'),
            GetString('YCO_SERVEURBI'),
            GetString('YCO_DATAMART'),false);
    monOlapCell.oxCube:=GetString('YCO_CUBE');
  end;
  except
  on E : Exception do
    begin
        PGIBOX(TraduireMemoire('Echec de connexion au cube : veuillez vérifier votre paramétrage !'));
        exit;
    end;
  end;

  if Hierarchy <> '' then
  begin
    monOlapCell.GetMembers(Hierarchy,1,'',ardetail);
      begin
        CMember.Clear;
        for iInd := vararraylowbound(ardetail,1) to vararrayhighbound(ardetail,1) do
        begin
          CMember.Items.Add(FormatteMember(ardetail[iInd]));
        end;
      end;
  end;
end;

function TOF_YYADMUSRBI.FormatteMember(Member:string):string;
var stemp,sresult:string;
begin
  result:='';
  stemp:=readtokenpipe(Member,'].[');
  while stemp <> '' do
  begin
    sresult:=stemp;
    stemp:=readtokenpipe(Member,'].[');
  end;
  result:=trim(StringReplace(sresult,']',' ',[rfReplaceAll]));
end;

procedure TOF_YYADMUSRBI.GFiltresOnClick (Sender : TObject) ;
var YFCTOB:TOB;
begin
  FiltreRow := TOB(GFiltres.Objects[0,GFiltres.Row]);
  if FiltreRow=nil then exit;
  SetControlText('EGUIDFILTRES',FiltreRow.GetString('YFL_GUID'));

  YFCTOB:=TOB.Create('YFILBASEEOLAP',nil,-1);
  YFCTOB.LoadDetailDB('YFILBASEEOLAP','"'+FiltreRow.GetString('YFL_GUID')+'"','',nil,False);
  if YFCTOB.Detail.Count = 0 then exit;
  CDimension.Text := YFCTOB.Detail[0].GetString('YFL_DIMENSION');
  CHierarchy.Text := YFCTOB.Detail[0].GetString('YFL_HIERARCHY');
  CMember.Text := YFCTOB.Detail[0].GetString('YFL_VALUE');
  FreeAndNil(YFCTOB);
end;

procedure TOF_YYADMUSRBI.CDimensionOnClick (Sender : TObject) ;
begin
  ChargeCHierarchy(CDimension.Text);
end;

procedure TOF_YYADMUSRBI.CHierarchyOnClick (Sender : TObject) ;
begin
  ChargeCMember(CHierarchy.Text);
end;


procedure TOF_YYADMUSRBI.BValidDroitOnClick(Sender:TObject);
var YUCTOB,YUCDETTOB,YBOTOB,YBODETTOB:TOB;
    S,St:String;
begin
  St:=CCube.text;
  S:=ReadTokenSt(St);

  YUCTOB:=TOB.Create('YUSRCUBEOLAP',nil,-1);
  if getControlText('EGUIDDROITS') = '' then
  begin
    YUCDETTOB:=TOB.Create('YUSRCUBEOLAP',YUCTOB,-1);
    YUCTOB.Detail[0].PutValue('YUC_GUID',AglGetGuid());
    YUCTOB.Detail[0].PutValue('YUC_USER',getControlText('ECODEUSR'));
  end
  else
  begin
    YUCTOB.LoadDetailDb('YUSRCUBEOLAP','"'+getControlText('EGUIDDROITS')+'"','',nil,False);
  end;

  YUCTOB.Detail[0].PutValue('YUC_CODCUBE',S);
  if CBAnonymous.Checked
  then
    YUCTOB.Detail[0].PutValue('YUC_ANONYMOUS','X')
  else
    YUCTOB.Detail[0].PutValue('YUC_ANONYMOUS','-');

  YUCTOB.Detail[0].PutValue('YUC_LOGINCWAS',GetControlText('ELOGIN'));
  YUCTOB.Detail[0].PutValue('YUC_PWDCWAS',GetControlText('EPASSWORD'));
  YUCTOB.Detail[0].PutValue('YUC_DOMCWAS',GetControlText('EDOMAIN'));
  YUCTOB.InsertOrUpdateDb;
  DisableCtrlsDroits;
  DisableCtrlsFiltres;

  AlimGDroits(getControlText('ECODEUSR'));

  FreeAndNil(YUCTOB);

  GDroitsOnClick(nil);
end;

procedure TOF_YYADMUSRBI.BInsertDroitOnClick(Sender:TObject);
Var QMn:TQuery;
begin
  VideCtrlsDroits;
  EnableCtrlsDroits;
  Qmn:= OpenSQL('SELECT YCO_CODCUBE, YCO_CUBE, YCO_LIBCUBE FROM YCUBEOLAP '+
    'WHERE YCO_CODCUBE NOT IN (SELECT YUC_CODCUBE FROM YUSRCUBEOLAP WHERE '+
    'YUC_USER = "'+getControlText('ECODEUSR')+'")',True);
  CCube.Clear;
  while not Qmn.Eof do
  begin
    CCube.Items.Add(QMn.Fields[0].AsString + ';' + QMn.Fields[1].AsString + ';' + QMn.Fields[2].AsString);
    Qmn.Next;
  end;

  Ferme(Qmn);

end;

procedure TOF_YYADMUSRBI.BDeleteDroitOnClick(Sender:TObject);
Var YUCTOB:TOB;
    stwhere:string;
    Qmn:TQuery;
begin
  YUCTOB:=TOB.Create('YUSRCUBEOLAP',nil,-1);
  stwhere := 'WHERE YUC_GUID="'+ getControlText('EGUIDDROITS')+'"';
  QMn :=  OpenSQL('SELECT * FROM YUSRCUBEOLAP ' + stWhere,True);
  try
    if not QMn.Eof then
        YUCTOB.LoadDetailDB('YUSRCUBEOLAP','','',QMn,true);
  finally
    Ferme(QMn);
  end;
  if (YUCTOB.Detail.Count > 0) then YUCTOB.Detail[0].DeleteDB;

  AlimGDroits(getControlText('ECODEUSR'));

  FreeAndNil(YUCTOB);
end;

procedure TOF_YYADMUSRBI.BDefaitDroitOnClick(Sender:TObject);
begin
  VideCtrlsDroits;
  DisableCtrlsDroits;
end;

procedure TOF_YYADMUSRBI.BValidFiltreOnClick(Sender:TObject);
Var YFATOB,YFADETTOB:TOB;
    stwhere:string;
    qmn:TQuery;
begin
  if not assigned(DroitRow) then exit;
  if CMember.Text='' then exit;
  if (TOBUSr.Detail.Count=0) then exit;

  YFATOB:=TOB.Create('YFILBASEOLAP',nil,-1);
  YFADETTOB:=TOB.Create('YFILBASEOLAP',YFATOB,-1);

  stwhere := 'WHERE YFL_DIMENSION="'+ CDimension.Text+'" AND'+
  ' YFL_HIERARCHY="'+ CHierarchy.Text +'" AND' +
  ' YFL_VALUE="'+ CMember.Text + '" AND' +
  ' YFL_YBOGUID="'+ getControlText('EGUIDBASE')+ '" AND'+
  ' YFL_USER="' + TOBUSR.Detail[0].GetString('US_UTILISATEUR') +'"';
  QMn :=  OpenSQL('SELECT * FROM YFILBASEOLAP ' + stWhere,True);
  try
    if QMn.Eof then
    begin
      YFATOB.Detail[0].PutValue('YFL_GUID',AglGetGuid());
      YFATOB.Detail[0].PutValue('YFL_YBOGUID',getControlText('EGUIDBASE'));
      YFATOB.Detail[0].PutValue('YFL_USER',TOBUSR.Detail[0].GetString('US_UTILISATEUR'));
      YFATOB.Detail[0].PutValue('YFL_DIMENSION',CDimension.Text);
      YFATOB.Detail[0].PutValue('YFL_HIERARCHY',CHierarchy.Text);
      YFATOB.Detail[0].PutValue('YFL_VALUE',CMember.Text);
      YFATOB.InsertOrUpdateDb;
    end
  finally
    Ferme(QMn);
  end;

  VideCtrlsFiltres;
  DisableCtrlsFiltres;

  AlimGFiltres(DroitRow.getString('YUC_USER'),DroitRow.getString('YUC_CODCUBE'));

  FreeAndNil(YFATOB);

  GFiltresOnClick(nil);

end;

procedure TOF_YYADMUSRBI.BInsertFiltreOnClick(Sender:TObject);
begin
  VideCtrlsFiltres;
  EnableCtrlsFiltres;
  if assigned(DroitRow) then ChargeCDimension(DroitRow.getString('YUC_CODCUBE'));
end;

procedure TOF_YYADMUSRBI.BDeleteFiltreOnClick(Sender:TObject);
Var YFATOB:TOB;
    stwhere:string;
    Qmn:TQuery;
begin
  YFATOB:=TOB.Create('YFILBASEOLAP',nil,-1);
  stwhere := 'WHERE YFL_GUID="'+ getControlText('EGUIDFILTRES')+'"';
  QMn :=  OpenSQL('SELECT * FROM YFILBASEOLAP ' + stWhere,True);
  try
    if not QMn.Eof then
        YFATOB.LoadDetailDB('YFILBASEOLAP','','',QMn,true);
  finally
    Ferme(QMn);
  end;
  if (YFATOB.Detail.Count > 0) then YFATOB.Detail[0].DeleteDB;

  GDroitsOnClick(nil);
  FreeAndNil(YFATOB);
end;

procedure TOF_YYADMUSRBI.BDefaitFiltreOnClick(Sender:TObject);
begin
  VideCtrlsFiltres;
  DisableCtrlsFiltres;
end;

procedure TOF_YYADMUSRBI.GDroitsOnDblClick (Sender : TObject) ;
begin
  EnableCtrlsDroits;
end;

procedure TOF_YYADMUSRBI.GFiltresOnDblClick (Sender : TObject) ;
begin
  EnableCtrlsFiltres;
end;

procedure TOF_YYADMUSRBI.CBAnonymousOnClick(Sender:TObject);
begin
  ELogin.Enabled := not(CBAnonymous.Checked);
  EPassWord.Enabled := not(CBAnonymous.Checked);
  EDomain.Enabled := not(CBAnonymous.Checked);
  if CBAnonymous.Checked then
  begin
    SetControlText('ELOGIN','');
    SetControlText('EPASSWORD','');
    SetControlText('EDOMAIN','');
  end;
end;

procedure TOF_YYADMUSRBI.BDroitGrpOnClick(Sender:TObject);
var UsersDuGrpTOB,YUCTOB,YUCDETAILTOB:TOB;
    stWhere:string;
    QMn:TQuery;
    iInd:integer;
begin
  if not assigned(DroitRow) then exit;

  if PGIASK('Voulez vous réellement appliquer le droit sur le cube ' +
                  DroitRow.getString('YCO_CUBE') +
                  ' à TOUS les utilisateurs du groupe ' +
                  TOBUSR.Detail[0].GetString('UG_LIBELLE') +
                  ' ?') = mrYes
    then
    begin
      UsersDuGrpTOB:=TOB.Create('UTILISAT',nil,-1);
      stwhere := 'WHERE US_GROUPE="'+TOBUSR.Detail[0].GetString('US_GROUPE')+'"'
        + ' AND US_UTILISATEUR <> "' + DroitRow.getString('YUC_USER') +'"';
      QMn :=  OpenSQL('SELECT * FROM UTILISAT ' + stWhere,True);
      try
        if not QMn.Eof then
            UsersDuGrpTOB.LoadDetailDB('LesUsersDeMonGroupe','','',QMn,true);
      finally
        Ferme(QMn);
      end;

      YUCTOB:=TOB.Create('YUSRCUBEOLAP',nil,-1);

      for iInd:=0 to UsersDuGrpTOB.Detail.Count - 1 do
      begin
        YUCDETAILTOB:=TOB.Create('YUSRCUBEOLAP',YUCTOB,-1);
        stwhere := 'WHERE YUC_USER="'+UsersDuGrpTOB.Detail[iInd].GetString('US_UTILISATEUR')+'"'
                + ' AND YUC_CODCUBE="' + DroitRow.getString('YUC_CODCUBE') +'"';
        QMn :=  OpenSQL('SELECT * FROM YUSRCUBEOLAP ' + stWhere,True);
        try
          if not QMn.Eof then
              YUCDETAILTOB.LoadDetailDB('YUSRCUBEOLAP','','',QMn,true)
          else
          begin
            YUCDETAILTOB.PutValue('YUC_GUID',AglGetGuid());
            YUCDETAILTOB.PutValue('YUC_USER',UsersDuGrpTOB.Detail[iInd].GetString('US_UTILISATEUR'));
            YUCDETAILTOB.PutValue('YUC_CODCUBE',DroitRow.GetString('YUC_CODCUBE'));
          end;
        finally
          Ferme(QMn);
        end;
        YUCDETAILTOB.PutValue('YUC_ANONYMOUS',DroitRow.GetString('YUC_ANONYMOUS'));
        YUCDETAILTOB.PutValue('YUC_LOGINCWAS',DroitRow.GetString('YUC_LOGINCWAS'));
        YUCDETAILTOB.PutValue('YUC_PWDCWAS',DroitRow.GetString('YUC_PWDCWAS'));
        YUCDETAILTOB.PutValue('YUC_DOMCWAS',DroitRow.GetString('YUC_DOMCWAS'));
      end;
      YUCTOB.InsertOrUpdateDb;
      FreeAndNil(YUCTOB);
    end;
    FreeAndNil(UsersDuGrpTOB);
end;

procedure TOF_YYADMUSRBI.BFiltreGrpOnClick(Sender:TObject);
var UsersDuGrpTOB,YFATOB,YFADETAILTOB:TOB;
    stWhere,styboguid:string;
    QMn:TQuery;
    iInd:integer;
begin
  if not assigned(DroitRow) then exit;
  if not assigned(FiltreRow) then exit;

  if PGIASK('Voulez vous réellement appliquer le filtre '+
                  FiltreRow.GetString('YFL_HIERARCHY') + ' / ' + FiltreRow.GetString('YFL_VALUE') +
                  ' sur la base ' +
                  DroitRow.GetString('YCO_DATAMART') +
                  ' à TOUS les utilisateurs du groupe ' +
                  TOBUSR.Detail[0].GetString('UG_LIBELLE') +
                  ' ?') = mrYes
      then
      begin
        UsersDuGrpTOB:=TOB.Create('UTILISAT',nil,-1);
        stwhere := 'WHERE US_GROUPE="'+TOBUSR.Detail[0].GetString('US_GROUPE')+'"'
          + ' AND US_UTILISATEUR <> "' + DroitRow.getString('YUC_USER') +'"';
        QMn :=  OpenSQL('SELECT * FROM UTILISAT ' + stWhere,True);
        try
          if not QMn.Eof then
              UsersDuGrpTOB.LoadDetailDB('LesUsersDeMonGroupe','','',QMn,true);
        finally
          Ferme(QMn);
        end;

        YFATOB:=TOB.Create('YFILBASEOLAP',nil,-1);

        for iInd:=0 to UsersDuGrpTOB.Detail.Count - 1 do
        begin
          stwhere := 'WHERE YCO_CODCUBE="' + DroitRow.getString('YUC_CODCUBE') +'"';
          QMn :=  OpenSQL('SELECT YCO_YBOGUID FROM YCUBEOLAP ' + stWhere,True);
          if Qmn.EOF
            then styboguid:=''
            else styboguid:= QMn.Fields[0].AsString;

          if styboguid<>'' then
          begin
            YFADETAILTOB:=TOB.Create('YFILBASEOLAP',YFATOB,-1);
            stwhere := 'WHERE YFL_YBOGUID="'+styboguid+'" AND YFL_USER ="'
            + UsersDuGrpTOB.Detail[iInd].getString('US_UTILISATEUR') +'"' ;
            QMn :=  OpenSQL('SELECT * FROM YFILBASEOLAP ' + stWhere,True);
            try
              if QMn.Eof then
              begin
                YFADETAILTOB.PutValue('YFL_GUID',AglGetGuid());
                YFADETAILTOB.PutValue('YFL_YBOGUID',styboguid);
                YFADETAILTOB.PutValue('YFL_USER',UsersDuGrpTOB.Detail[iInd].getString('US_UTILISATEUR'));
                YFADETAILTOB.PutValue('YFL_DIMENSION',FiltreRow.GetString('YFL_DIMENSION'));
                YFADETAILTOB.PutValue('YFL_HIERARCHY',FiltreRow.GetString('YFL_HIERARCHY'));
                YFADETAILTOB.PutValue('YFL_VALUE',FiltreRow.GetString('YFL_VALUE'));
              end;
            finally
              Ferme(QMn);
            end;
          end;
        end;
        YFATOB.InsertOrUpdateDb;
        FreeAndNil(YFATOB);
      end;
      FreeAndNil(UsersDuGrpTOB);
end;

Initialization
  registerclasses ( [ TOF_YYADMUSRBI ] ) ;
end.
