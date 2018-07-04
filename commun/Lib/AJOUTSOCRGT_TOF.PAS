{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 29/12/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AJOUTSOCRGT ()
Mots clefs ... : TOF;AJOUTSOCRGT
*****************************************************************}
Unit AJOUTSOCRGT_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
     uTob,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     UTOB,
     HsysMenu,
     HTB97;

Type
  TOF_AJOUTSOCRGT = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    CodeMultiDossier:string;
    BValider: TToolbarButton97;
    HMTrad: THSystemMenu;
    CBMultiSoc : THValComboBox;
    TLstBases,TLstUsers : TListBox;
    GBaseSupp : THGrid;
    RgtPrincTOB,RgtUserTob,PointIniTOB,BasesRgtTOB,BasesRgtDetailTOB,AddBaseTOB,AddBaseDetailTOB:TOB;
    procedure EnregSaisie;
    procedure CbMultiSocOnClick (Sender : TObject);
    procedure GBaseSuppOnClick (Sender : TObject);
  end ;

Implementation

uses utilpgi, umultidossierutil, utobdebug, yybundle_tof,majhalleyutil,wcommuns,
     USatUtil;  //Pour IsOnlyUser()

Const
  { code du bundle Regroupements dossiers' }
  sCodeRDoss = 'YS3';


procedure TOF_AJOUTSOCRGT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AJOUTSOCRGT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AJOUTSOCRGT.OnUpdate ;
var iInd, iIndTob:integer;
    T:TOB;
begin
  Inherited ;
  if GBaseSupp.nbSelected = 0 then
  begin
    PGIBox('Vous devez sélectionner au moins une société');
    SetFocusControl('GBASESUPP');
    exit;
  end;

  {LAU le 10/12/2007 Correction pour le Client TESSI pour permettre de traiter plusieurs bases sélectionnées
  if GBaseSupp.nbSelected > 1 then
  begin
    PGIBox('Vous ne devez sélectionner qu''une société');
    SetFocusControl('GBASESUPP');
    exit;
  end;}

  //Tester que personne ne travaille sur la base.
  if not IsOnlyUser('ATTENTION') then exit;

  if PGIAsk('Confirmez-vous le rajout de la (des) base(s) sélectionnée(s)'+
        ' dans le regroupement '+ Cbmultisoc.Value  +' ?' + chr(13) +
        'ATTENTION, cette opération est IRREVERSIBLE et n''est destinée '+
        'qu''à l''ajout de sociétés n''ayant pas encore démarré leur activité.') <> mrYes then exit;

  for iInd := 1 to GBaseSupp.RowCount -1 do
    if GBaseSupp.IsSelected(iInd) and (GBaseSupp.Cells[0,iInd] <> '') then
      begin
        T := AddBaseTOB.FindFirst(['DOSSIER'],[GBaseSupp.Cells[0,iInd]],false);
        if T=nil then exit;
//        PGIBOX('OK Rajout '+ GBaseSupp.Cells[0,iInd]);
        BasesRgtDetailTOB:=TOB.Create('',BasesRgtTOB,-1);
        iIndTob:=BasesRgtTOB.Detail.Count - 1;
        BasesRgtTOB.Detail[iIndTob].AddChampSupValeur('BASE',T.GetString('BASE'));
        BasesRgtTOB.Detail[iIndTob].AddChampSupValeur('DOSSIER',T.GetString('DOSSIER'));
        FreeandNil(T);
      end;

  EnregSaisie;
  SetControlEnabled('BVALIDER',false);

end ;

procedure TOF_AJOUTSOCRGT.EnregSaisie ;
var iInd : integer;
    st:string;
    TobBundle,TobB:TOB;
    ValideBundle:tValideBundle;

  function EstBundlePartage(sCode : string) : boolean;
  var sBaseMs,UneBase : string;
  begin
    result := (sCode = sCodeRDoss);
    if not result then
    begin
      sBaseMs := GetBasesMS('', true);
      UneBase := ReadTokenSt (sBaseMs) ;
      while (UneBase <> '') and (not result) do
      begin
        result := not ExisteSql('SELECT CO_CODE FROM COMMUN WHERE CO_TYPE="YEB" '
              + 'AND CO_LIBRE="' + sCode + '" AND CO_ABREGE<>"BUN" AND CO_LIBELLE NOT IN'
              + '(SELECT IIF(DS_VUE="",DS_NOMTABLE,DS_NOMTABLE||";"||DS_VUE) AS CO_LIBELLE '
              + 'FROM ' + GetBase (UneBase, 'DESHARE')+')');
        UneBase := ReadTokenSt (sBaseMs) ;
      end;
    end;
  end;

begin
  st := '';

  for iInd := 0 to BasesRgtTOB.Detail.Count - 1 do
    st := st +
      BasesRgtTOB.Detail[iInd].GetValue('DOSSIER') +
      '|' +
      BasesRgtTOB.Detail[iInd].GetValue('BASE') +
      ';';

  st := st + #13#10 + RgtUserTob.Detail[0].GetString('USERGRP');

  try
    RgtUserTob.Detail[0].PutValue('YMD_DATEMODIF', Now);
    RgtUserTob.Detail[0].PutValue('YMD_UTILISATEUR', V_PGI.User);
    RgtUserTob.Detail[0].PutValue('YMD_DETAILS', st);
    RgtUserTob.UpdateDB;

  //!!!!!!!!!!!!!!!!!!!!ATTENTION : METTRE A JOUR ##MULTISOC
    RgtPrincTOB.Detail[0].PutValue('YMD_DATEMODIF', RgtUserTob.Detail[0].getString('YMD_DATEMODIF'));
    RgtPrincTOB.Detail[0].PutValue('YMD_UTILISATEUR', V_PGI.User);
    RgtPrincTOB.Detail[0].PutValue('YMD_DETAILS', st);
    RgtPrincTOB.UpdateDB;

{  //!!!!!!!!!!!!!!!!!!!!ATTENTION : METTRE A JOUR DOSSIER
    for iInd := 0 to BasesRgtTOB.Detail.Count - 1 do
    begin
      UneBase:=BasesRgtTOB.Detail[iInd].getString('BASE');
      UpdateMultiDossier (UneBase, false);
    end;

  //  wInitProgressForm(nil, TraduireMemoire('Enregistrement du paramétrage'), Ecran.Caption, 6, False, True);
  //  Transactions (ValideParametrage, 1);
  //  wFiniProgressForm;

    if V_PGI.IoError <> oeOk then
      PgiBox ('Erreur lors de l''enregistrement du paramétrage')
    else PgiBox ('Le paramétrage a bien été enregistré'#13'Veuillez re-valider votre partage et vous connecter à nouveau'#13'pour prendre en compte le paramétrage');

    YYLanceFiche_Bundle('YY', 'YYBUNDLE', '', '', '');}

  TobBundle:=TOB.Create('',nil,-1);
  try
    TobBundle.LoadDetailDBFromSQL ('COMMUN',
       'SELECT CO_CODE, CO_LIBELLE, CO_ABREGE FROM COMMUN WHERE CO_TYPE="YBU" ORDER BY CO_ABREGE');

    for iInd := 0 to TobBundle.Detail.Count -1 do
    begin
      TobB := TobBundle.Detail[iInd];
      if EstBundlePartage(TobB.GetString ('CO_CODE')) then
        TobB.AddChampSupValeur ('PARTAGE', 'X' , false)
      else
        TobB.AddChampSupValeur ('PARTAGE', '-' , false);
    end;

    ValideBundle:= tValideBundle.Create(CodeMultiDossier,TobBundle);
    try
      ValideBundle.ValideParametrage;
    finally
      begin
        PgiBox ('Le paramétrage a bien été enregistré'#13'Veuillez vous connecter à nouveau pour prendre en compte le paramétrage');
        ValideBundle.Free;
      end;
    end;
  finally
    if assigned (TobBundle) then FreeAndNil(TobBundle);
  end;
  except on E: Exception do
      PGiError(E.ClassName + ': ' + E.Message);
  end;

  THValCombobox(GetControl('CBMULTISOC')).Enabled:=false;
  TListBox(GetControl('LSTBASES')).Enabled:=false;
  TListBox(GetControl('LSTUSERS')).Enabled:=false;
  THGrid(GetControl('GBASESUPP')).Enabled:=false;
  TToolBarButton97(GetControl('BVALIDER')).Enabled:=false;

end;

procedure TOF_AJOUTSOCRGT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AJOUTSOCRGT.OnArgument (S : String ) ;
var stWhere  :string;
    Qmn:TQuery;
    lQMD     : TQuery ;
    lTSData  : TStringList ;
    lStVal   : String ;
begin
  Inherited ;
  if not IsMonoOuCommune then
  begin
    PGIBOX(TraduireMemoire('Cette fonctionnalité n''est autorisée que sur la base commune d''un regroupement.'));
    exit;
  end;

// récupération paramétrage du regroupement
  lStVal := '' ;
  lQMD   := OpenSQL('SELECT * FROM YMULTIDOSSIER WHERE YMD_CODE = "' + MS_CODEREGROUPEMENT + '"', True ) ;
  if not lQMD.Eof then
    lStVal := lQMD.FindField('YMD_DETAILS').AsString ;
  Ferme( lQMD ) ;
  if lStVal = '' then Exit ;

// Récupération 1ère ligne : on prendra ensuite les groupes de user
  lTSData      := TStringList.Create ;
  lTSData.Text := lStVal ;
  lStVal       := lTSData.Strings[0] ;

  RgtPrincTOB:=TOB.Create('YMULTIDOSSIER',nil,-1);
  stwhere := 'WHERE YMD_CODE="'+MS_CODEREGROUPEMENT+'"';
  QMn :=  OpenSQL('SELECT * FROM YMULTIDOSSIER ' + stWhere,True);
  try
    if not QMn.Eof then RgtPrincTOB.LoadDetailDB('YMULTIDOSSIER','','',QMn,true);
  finally
    Ferme(QMn);
  end;

  if assigned(GetControl('CBMULTISOC')) then
  begin
    CbMultiSoc := THValCombobox(GetControl('CBMULTISOC'));
    CbMultiSoc.OnClick:=CbMultiSocOnClick;
    CbMultiSoc.Plus := ' YMD_CODE<>"' + MS_CODEREGROUPEMENT +
      '"' + ' AND YMD_DETAILS LIKE "' + lstVal + '%"';
  end;
  if assigned(GetControl('LSTBASES')) then
    TLstBases := TListBox(GetControl('LSTBASES'));
  if assigned(GetControl('LSTUSERS')) then
    TLstUsers := TListBox(GetControl('LSTUSERS'));

  if assigned(GetControl('GBASESUPP')) then
  begin
    GBaseSupp := THGrid(GetControl('GBASESUPP'));
    GBaseSupp.OnClick:=GBaseSuppOnClick;
  end;
  AddBaseTOB:=TOB.Create('',nil,-1);

  if assigned(GetControl('BVALIDER')) then
  begin
    BValider := TToolbarButton97(GetControl('BVALIDER'));
  end;


end ;

procedure TOF_AJOUTSOCRGT.OnClose ;
begin
  Inherited ;
  FreeAndNil(AddBaseTOB);
  FreeAndNil(PointIniTOB);
  FreeAndNil(BasesRgtTOB);
  FreeAndNil(RgtUserTob);
  FreeAndNil(RgtPrincTOB);
end ;

procedure TOF_AJOUTSOCRGT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AJOUTSOCRGT.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_AJOUTSOCRGT.CbMultiSocOnClick (Sender : TObject);
var
  stLesBases, UneBase , UnUser, LaBase, LeDossier, StWhere : string;
  iInd,iIndTob:integer;
  T:TOB;
  Qmn:TQuery;
  lQMD     : TQuery ;
  lTSData  : TStringList ;
  lStVal   : String ;
  lStBase  : String ;
  lStUsers   : String ;
  intVersionRef,intVersionBase: integer;
begin
  TLstBases.Clear;
  TLstUsers.Clear;
  AddBaseTOB.ClearDetail;

  CodeMultiDossier := ThValComboBox(Sender).Value;

  // récupération paramétrage du regroupement
  lStVal := '' ;
  lQMD   := OpenSQL('SELECT * FROM YMULTIDOSSIER WHERE YMD_CODE = "' + CodeMultiDossier + '"', True ) ;
  if not lQMD.Eof then
    lStVal := lQMD.FindField('YMD_DETAILS').AsString ;
  Ferme( lQMD ) ;
  if lStVal = '' then Exit ;

// Récupération 1ère ligne : on prendra ensuite les groupes de user
  lTSData      := TStringList.Create ;
  lTSData.Text := lStVal ;
  lStVal       := lTSData.Strings[0] ;
  lStUsers     := lTSData.Strings[1] ;

  while lStVal<>'' do
    begin
    lStBase := ReadTokenSt( lStVal ) ;
    stLesBases := stLesBases + lStBase + ';';
    end ;

  if Pos (V_PGI.SchemaName, stLesBases) = 0 then
  begin
    PGIBox ('La base connectée ne fait pas partie du regroupement');
    TLstBases.Clear;
    TLstUsers.Clear;
    exit;
  end
  else
  begin
    RgtUserTob := TOB.Create('YMULTIDOSSIER', nil, -1);
    stwhere := 'WHERE YMD_CODE="'+CodeMultidossier+'"';
    QMn :=  OpenSQL('SELECT * FROM YMULTIDOSSIER ' + stWhere,True);
    try
      if not QMn.Eof then
        begin
          RgtUserTob.LoadDetailDB('YMULTIDOSSIER','','',QMn,true);
          RgtUserTob.Detail[0].AddChampSupValeur('USERGRP',lStUsers);
        end;
    finally
      Ferme(QMn);
    end;

    BasesRgtTOB:=TOB.Create('LesBases',nil,-1);
    UneBase:=ReadTokenSt(stLesBases);
    iInd:=0;
    while UneBase <> '' do
    begin
      BasesRgtDetailTOB:=TOB.Create('',BasesRgtTOB,-1);
      BasesRgtTOB.Detail[iInd].AddChampSupValeur('DOSSIER',ReadTokenPipe(UneBase,'|'));
      TLstBases.Items.Add(BasesRgtTOB.Detail[iInd].getString('DOSSIER'));
      BasesRgtTOB.Detail[iInd].AddChampSupValeur('BASE',UneBase);
      UneBase:=ReadTokenSt(stLesBases);
      iInd:=iInd+1;
    end;

    UnUser:=ReadTokenSt(lStUsers);
    while UnUser <> '' do
    begin
      TLstUsers.Items.Add(UnUser);
      UnUser:=ReadTokenSt(lStUsers);
    end;
  end;

  intVersionRef := VersionBase(V_PGI.CurrentAlias); 

  PointIniTOB:=TOB.Create('',nil,-1);
  ChargeMultiDossiersFromIni(PointIniTOB);
  iIndTob:=0;
  wInitProgressForm(nil, TraduireMemoire('Construction de la liste des bases disponibles'), Ecran.Caption, 6, False, True);
  for iInd := 1 to PointIniTOB.Detail.Count - 1 do
  begin
    LaBase := PointIniTOB.Detail[iInd].getString('BASE');
    LeDossier := PointIniTOB.Detail[iInd].getString('DOSSIER');
    intVersionBase:=VersionBase(LeDossier);
    wMoveProgressForm ('Base '+LeDossier+ ' en version '+ IntToStr(intVersionBase));
    T := BasesRgtTOB.FindFirst(['BASE'],[LaBase],false);
    if (T = nil) and (intVersionBase=intVersionRef) then
    begin
      AddBaseDetailTOB:=TOB.Create('',AddBaseTOB,-1);
      AddBaseTOB.Detail[iIndTob].AddChampSupValeur('BASE',LaBase);
      AddBaseTOB.Detail[iIndTob].AddChampSupValeur('DOSSIER',LeDossier);
      iIndTob:=iIndTob+1;
    end;
//    FreeAndNil(T); !!!!!!!!!!!!!!!ATTENTION : en fait vire l'enreg trouve par le findfirst
  end;                                                                          
  wFiniProgressForm;
  AddBaseTOB.PutGridDetail (GBaseSupp, False, False, 'DOSSIER', True);

  HMTrad.ResizeGridColumns(GBaseSupp);
end;

procedure TOF_AJOUTSOCRGT.GBaseSuppOnClick (Sender : TObject);
begin
  GBaseSupp.FlipSelection(GBaseSupp.Row);
end;


Initialization
  registerclasses ( [ TOF_AJOUTSOCRGT ] ) ;
end.
