{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 09/08/2006
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : DPDTPROCESS (DPDTPROCESS)
Mots clefs ... : TOM;DPDTPROCESS
*****************************************************************}
Unit dpTOM_DPDtProcess ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fiche,
     FichList,
{$else}
     eFiche,
     eFichList,
{$ENDIF}
{$IFDEF VER150}
     Variants,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOM,
     UTob ;

Type
  TMonProc = procedure (sender:TObject) ;
  TOM_DPDTPROCESS = Class (TOM)

  private
    UneModif, ModifAnnexe:boolean ;
    LstDossierIdemIni : TStringList ;
    procDeleteDecla : TNotifyEvent ;
    IdPourDupli : string ;
    AValider : boolean ;
    procValideDecla : TNotifyEvent ;
    procedure DDP_MetierClick (Sender: TObject);
    procedure bValider_OnClick(sender:TObject) ;
    procedure bRetireClick (Sender: TObject);
    procedure bAjouteClick (Sender: TObject);
    procedure popDeleteOneClick (Sender: TObject);
    procedure popDeleteAllClick (Sender: TObject);
    procedure setLocalEvents ;
    function  getCompteurTable : integer ;
    function  IsEquivalent (NoDossier : string; ExcluTestId : boolean = false) : boolean ;
    function  _EnregOk : boolean ;
    procedure LoadDossierIdentique ;
    function  LstDossier2SqlIN (laLst:string) : string ;
    procedure SauveDossierIdentique ;
    function  getWhereDossierIdentique : string;
    procedure TraiteDossierAjouteSupprime ;

  public
    procedure OnNewRecord ; override ;
    procedure OnUpdateRecord ; override ;
    procedure OnLoadRecord ; override ;
    procedure OnChangeField (F : TField) ; override ;
    procedure OnArgument (stArgument : String ) ; override ;
    procedure OnClose                    ; override ;

  end ;

implementation

uses hdb,
  {$IFNDEF EAGLCLIENT} //+LMO20060901
     fe_main,
  {$ELSE}
     MaineAGL,
  {$ENDIF}             //-LMO20060901
  uSATUtil, menus, htb97, DpJurOutils;

procedure TOM_DPDTPROCESS.OnNewRecord ;
var DDP_ : Tob ;
begin
  inherited ;
  if IdPourDupli<>'' then
  begin
    DDP_ := Tob.create('DPDTPROCESS', nil, -1) ;
    if DDP_.selectDB(IdPourDupli, nil) then
    begin
      DDP_.p('DDP_DPDTPROCESS', 0) ;
      DDP_.PutEcran(TFFiche(Ecran)) ;
    end ;
    DDP_.free ;
  end
  else
  begin
    setControlText('DDP_NODOSSIER','') ;
    setControlText('DDP_DATEDU',DateTostr(idate1900)) ;
    setControlText('DDP_DATEAU',DateTostr(idate2099)) ;
    setControlText('LBLNODOSSIER','') ;
    setControlText('LBLFINALITE','') ;
  end ;
end ;

function TOM_DPDTPROCESS.IsEquivalent (NoDossier : string; ExcluTestId : boolean = false) : boolean ;
var q : TQuery ;
    st : string ;
begin
  st :=   'select count(*) from DPDTPROCESS ' +
          'where DDP_NODOSSIER="' + NoDossier + '" '+
          'and DDP_UTILISATEUR="'+GetControlText('DDP_UTILISATEUR')+'" '+
          'and DDP_FINALITE="'+GetControlText('DDP_FINALITE')+'" ' +
          'and DDP_MISSION="'+GetControlText('DDP_MISSION')+'" ' ;
  if not ExcluTestId then st := st + 'and DDP_DPDTPROCESS<>'+GetControlText('DDP_DPDTPROCESS') ;
  q:=openSql(st,true) ;
  result := q.Fields[0].asInteger > 0 ;
  ferme (q) ;
end ;

function TOM_DPDTPROCESS._EnregOk : boolean ;

    function NonVide (ch:string) : boolean ;
    begin
      result:=false ;
      if GetControlText (ch)='' then
      begin
        //PgiInfo(ChampToLibelle(ch) ' doit être renseigné.') ;
        LastError:=1 ;
        LastErrorMsg := ChampToLibelle(ch) + ' doit être renseigné.' ;
        FocusControle(GetControl('pGeneral').FindComponent(ch), true);
        exit ;
      end ;
      result:=true ;
    end ;
begin
  result:=false ;
  if not (NonVide('DDP_NODOSSIER') or NonVide('DDP_UTILISATEUR')
   or NonVide('DDP_FINALITE') or NonVide('DDP_MISSION')) then exit ;

  if IsEquivalent (GetControlText('DDP_NODOSSIER')) then
  begin
    LastError:=2 ;
    LastErrorMsg := 'Cette fiche existe déjà.' ;
    exit ;
  end ;

  result := true ;
end ;

function  TOM_DPDTPROCESS.getCompteurTable : integer ;
var q : tQuery;
    n : integer ;
begin
    q:=openSql('select max(DDP_DPDTPROCESS) from DPDTPROCESS',True);
    n := q.Fields[0].asInteger ;
    ferme(q) ;
    result := n+1 ;
end ;

procedure TOM_DPDTPROCESS.OnUpdateRecord ;
var n : integer ;
    ddp_ : Tob ;
begin
  inherited ;

  if not _EnregOk then exit ;

  if TFFiche(Ecran).TypeAction<>taCreat then
  begin
    SauveDossierIdentique ;
    ddp_:= tob.create('DPDTPROCESS', nil, -1) ;
    for n:=1 to ddp_.NbChamps do
      if isfieldModified( ddp_.GetNomChamp(n) ) then begin UneModif:=true; break; end ;
    ddp_.free;
  end
  else
    UneModif:=true ;

  TraiteDossierAjouteSupprime ;
  if GetField('DDP_DPDTPROCESS')=0 then SetField('DDP_DPDTPROCESS', getCompteurTable) ;

end ;

procedure TOM_DPDTPROCESS.OnChangeField (F : TField) ;
begin
  inherited ;
end ;

procedure TOM_DPDTPROCESS.OnLoadRecord ;
begin
  inherited ;
  UneModif := false ;
  ModifAnnexe := false ;
  if TFFiche(Ecran).TypeAction<>taCreat then LoadDossierIdentique ;//Doublon avec onNewRecord?
end ;

procedure TOM_DPDTPROCESS.OnArgument (stArgument : String ) ;
var i : integer ;
begin
  inherited ;
  setLocalEvents;
  LstDossierIdemIni:=TStringList.create ;
  if Pos('DUPLI', stArgument)>0 then
  begin
    i := postfs(stArgument, ';', 'DUPLI') ;
    if i>0 then idPourDupli := gtfs(stArgument, ';', i+1) ;
  end ;
end ;

procedure TOM_DPDTPROCESS.OnClose ;
begin
  Inherited ;
  if AValider and (UneModif or ModifAnnexe) and (pos('REFRESH;', TFFiche(Ecran).Retour)=0) then
    TFFiche(Ecran).Retour := TFFiche(Ecran).Retour + 'REFRESH;' ;
  if AValider and ModifAnnexe then TraiteDossierAjouteSupprime ;
  TFFiche(Ecran).Retour := TFFiche(Ecran).Retour + 'DDP_DPDTPROCESS;' +GetControlText('DDP_DPDTPROCESS');
  LstDossierIdemIni.Free ;
  if TFFiche(Ecran).TypeAction<>taCreat then TFFiche(Ecran).Close ;
end ;


procedure TOM_DPDTPROCESS.setLocalEvents ;
begin
  THEdit(GetControl('DDP_Metier')).OnClick := DDP_MetierClick;
  THEdit(GetControl('BRETIRE')).OnClick := bRetireClick;
  THEdit(GetControl('BAJOUTE')).OnClick := bAjouteClick;

  if assigned(TToolBarButton97(GetControl('BVALIDER')).OnClick) then
    procValideDecla := TToolBarButton97(GetControl('BVALIDER')).OnClick ;
  TToolBarButton97(GetControl('BVALIDER')).OnClick := bValider_OnClick;

  if assigned(TToolBarButton97(GetControl('BDELETE')).OnClick) then
    procDeleteDecla := TToolBarButton97(GetControl('BDELETE')).OnClick ;
  TToolBarButton97(GetControl('BDELETE')).OnClick := nil ;
  TMenuItem(GetControl('POPDELETEALL')).OnClick  := popDeleteAllClick ;
  TMenuItem(GetControl('POPDELETEONE')).OnClick  := popDeleteOneClick ;

end ;


procedure TOM_DPDTPROCESS.DDP_MetierClick(Sender: TObject);
var finalite : thDbEdit;
    Metier : thValComboBox ;
begin
  finalite := thDbEdit(getControl('DDP_FINALITE')) ;
  Metier := thValComboBox(getControl('DDP_METIER')) ;

  if (finalite=nil) or (Metier=nil) then exit ;

  finalite.plus := ' and CO_LIBRE = "' + getControlText('DDP_METIER') + '" ' ;
end ;

function TOM_DPDTPROCESS.LstDossier2SqlIN (laLst:string) : string ;
var i : integer ;
    lst : TListBox ;
    st : string ;
begin
  lst := TListBox(getcontrol(laLst)) ;
  for i:= 0 to lst.Items.Count - 1 do
  begin
    if st<>'' then st := st + ',' ;
    st := st + '"' + gtfs(lst.Items[i],' -',1) + '"';
  end ;
  if st<>'' then st := '(' + st + ')' ;
  result := st ;
end ;


function vString(v:variant):String;//LMO20060901 dans EntDP?
var st:string;
begin
  if VarIsNull(v) then begin result:='' ; exit ; end ;
  st:=VarAsType(v,VarOleStr);
  if St=#0 then St:='';
  Result:=St;
end;

function vDate(v:variant):TDateTime; //LMO20060901 dans EntDP?
var d:TDateTime;
begin
try
   if VarIsNull(v) then begin result:=iDate1900 ; exit ; end ;
   d:=VarAsType(v,VarDate);
   if d<iDate1900 then Result:=iDate1900
                  else Result:=d;
except
   result:=iDate1900;
   end ;
end;

function vDouble(v:variant):Double;//LMO20060901 dans EntDP?
var st:string ;
begin
if VarType(v) in [varSmallint,varInteger,varSingle,varDouble,varCurrency]
   then result:=VarAsType(V,VarDouble)  else
    BEGIN
    St:=vstring(v) ;
    if IsNumeric(St) then result:=Valeur(St) else result:=0 ;
    END ;
end;

function TOM_DPDTPROCESS.getWhereDossierIdentique : string ;     //LMO20060901

    function _date (Fld:string): string ;
    begin
    result := 'and ' + Fld + '="' + UsDateTime(vDate(BufferAvantModif.g(Fld))) + '" '  ;
    end ;
begin

  result := ' DDP_UTILISATEUR="' + vString(BufferAvantModif.g('DDP_UTILISATEUR')) + '" ' +
            'and DDP_METIER="' + vString(BufferAvantModif.g('DDP_METIER')) + '" ' +
            'and DDP_FINALITE="' + vString(BufferAvantModif.g('DDP_FINALITE')) + '" ' +
            'and DDP_MISSION="' + vString(BufferAvantModif.g('DDP_MISSION')) + '" ' +
            'and DDP_SEUIL=' + StrfPoint(vDouble(BufferAvantModif.g('DDP_SEUIL'))) + ' ' +
            'and DDP_NODOSSIER <> "' + vString(BufferavantModif.g('DDP_NODOSSIER')) + '" ' +
            _date('DDP_DATEDU') + _date ('DDP_DATEAU') + ' ' ;

end ;

procedure TOM_DPDTPROCESS.LoadDossierIdentique ;
var q : TQuery ;
    lb : TListBox ;
begin
  q:=openSql('select DDP_NODOSSIER, DOS_LIBELLE from DPDTPROCESS P ' +
             'left join DOSSIER D on D.DOS_NODOSSIER = P.DDP_NODOSSIER ' +
             ' where ' + getWhereDossierIdentique , true, -1, '', True);
  lb := TListBox(getControl('LstDossiers')) ;
  if lb <> nil then
  begin
    lb.Items.Clear ;
    while not q.eof do
    begin
      lb.Items.Add(q.FindField('DDP_NODOSSIER').asString + ' - ' + q.FindField('DOS_LIBELLE').asString) ;
      LstDossierIdemIni.Add(q.FindField('DDP_NODOSSIER').asString) ;
      q.Next ;
    end ;
  end ;
  Ferme(q) ;
end ;

procedure TOM_DPDTPROCESS.SauveDossierIdentique ;
var st, Dossier : string ;
    n : integer ;
begin
  if TListBox(GetControl('LstDossiers')).Items.Count=0 then exit ;

  if pgiAsk('Souhaitez-vous reporter les modifications dans '#13 +
            'la liste des dossiers ayant un paramétrage identique?')<> mrYes then exit ;

  st := 'update DPDTPROCESS set ' +
        'DDP_UTILISATEUR ="' + getControlText('DDP_UTILISATEUR') + '", ' +
        'DDP_METIER ="' + getControlText('DDP_METIER') + '", ' +
        'DDP_FINALITE ="' + getControlText('DDP_FINALITE') + '", ' +
        'DDP_MISSION ="' + getControlText('DDP_MISSION') + '", ' +
        'DDP_SEUIL =' + StrfPoint(strToFloat(getControlText('DDP_SEUIL'))) + ', ' +
        //'DDP_NODOSSIER = "' + getControlText('DDP_NODOSSIER') + '", ' +
        'DDP_DATEDU = "' + usdateTime(strToDate(getControlText('DDP_DATEDU'))) +'", ' +
        'DDP_DATEAU = "' + usdateTime(strToDate(getControlText('DDP_DATEAU'))) +'" ' +
        'where ' + getWhereDossierIdentique  ;
  Dossier := LstDossier2SqlIN('LstDossiers');
  if Dossier <> '' then st := st + ' and DDP_NODOSSIER in ' + Dossier ;
  n := ExecuteSql(st) ;
end ;

procedure TOM_DPDTPROCESS.bValider_OnClick(sender:TObject) ;
begin
  AValider := true ;
  procValideDecla (sender) ;
end ;

procedure TOM_DPDTPROCESS.bRetireClick (Sender: TObject);
var lst, lstD, lstC : TListBox ;
    i : integer ;
    st : string ;
begin
  lst := TListBox(getcontrol('LstDossiers')) ;
  lstD := TListBox(getcontrol('LstDel')) ;
  lstC := TListBox(getcontrol('LstCreate')) ;

  if lst.ItemIndex>-1 then
  begin
    st := gtfs(lst.Items[lst.ItemIndex],' -',1) ;
    i:= lstC.Items.IndexOf(st);
    if i>-1 then lstC.Items.delete(i) ;
    i:= lstD.Items.IndexOf(st);
    if (i=-1) and (LstDossierIdemIni.IndexOf(st)>-1) then lstD.Items.Add(st) ;
    lst.Items.delete(lst.ItemIndex) ;
    ModifAnnexe:=true ;
    //ModeEdition(DS) ; //SetField('DDP_NODOSSIER', GetField('DDP_NODOSSIER')) ;
  end
  else
    PgiInfo('Aucun dossier n''a été sélectionné dans la liste des autres dossiers.');
end ;

procedure TOM_DPDTPROCESS.bAjouteClick (Sender: TObject);
var lst, lstD, lstC : TListBox ;
    LstErr, res, st, code, lib : string ;
    i, n : integer ;
begin
    res := AglLanceFiche('YY', 'YYDOSSIER_SEL', '', '', ';MULTISELECT;') ;
    if res = '' then exit ;
    lst := TListBox(getcontrol('LstDossiers')) ;
    lstD := TListBox(getcontrol('LstDel')) ;
    lstC := TListBox(getcontrol('LstCreate')) ;

    n:=1; st:='*';
    while st<>'' do
    begin
      st:=gtfs(res,'|', n);
      if st<>'' then
      begin
        Code := gtfs(st,';', 1) ;
        Lib := gtfs(st,';', 3) ;

        if IsEquivalent(Code, true) or (Code = GetControlText('DDP_NODOSSIER')) then
        begin
          //PgiInfo('Cette fiche existe déjà.') ;
          //exit ;
          LstErr := LstErr + 'Cette fiche existe déjà : ' + Code + ' - ' + Lib + #13 ;
        end
        else
        begin
          i := lst.items.indexOf(st) ;
          if i=-1 then
          begin
            lst.ItemIndex := lst.Items.add(Code + ' - ' + Lib) ;
            i:= lstC.Items.IndexOf(Code); if i=-1 then lstC.Items.Add(Code) ;
            i:= lstD.Items.IndexOf(Code); if i>-1 then lstD.Items.Delete(i) ;
            ModifAnnexe:=true ;
          end
          else
          begin
            //PgiInfo('Le dossier ' + Lib + ' est déjà présent dans la liste.');
            LstErr := LstErr + 'Le dossier ' + Code + ' - ' + Lib + ' est déjà présent dans la liste.' ;
            lst.ItemIndex := i ;
          end ;
        end ;
      end ;
      inc(n) ;
  end ;
  if LstErr<>'' then PgiInfo(LstErr);
end ;

procedure TOM_DPDTPROCESS.popDeleteOneClick (Sender: TObject);
begin
  TFFiche(Ecran).Retour := 'REFRESH;';
  procDeleteDecla(sender);
end ;

procedure TOM_DPDTPROCESS.popDeleteAllClick (Sender: TObject);
var st, dossier : string ;
    n : integer ;
begin

  if pgiask('Confirmez-vous la suppression de ces enregistrements?')<>mrYes then exit ;
  st := 'delete DPDTPROCESS ' +
        'where (' + getWhereDossierIdentique  ;
  Dossier := LstDossier2SqlIN('LstDossiers');
  if Dossier <> '' then st := st + ' and DDP_NODOSSIER in ' + Dossier ;
  st := st + ') or DDP_DPDTPROCESS='+GetControlText('DDP_DPDTPROCESS') ;
  n:=executeSql(st) ;
  if n>1 then PgiBox(intTostr(n) + ' enregistrements supprimés.')
         else PgiBox(intTostr(n) + ' enregistrement supprimé.') ;
  TFFiche(Ecran).Retour := 'REFRESH;';
  TFFiche(Ecran).Close ;
end ;

procedure TOM_DPDTPROCESS.TraiteDossierAjouteSupprime ;
var st, dossier : string ;
    i, n : integer ;
    LstC : TListBox;

begin
  //dossier idem supprimé
  Dossier := LstDossier2SqlIN('LstDel');
  if Dossier <> '' then
  begin
    st := 'delete DPDTPROCESS ' +
          'where (' + getWhereDossierIdentique + ') and DDP_NODOSSIER in ' + Dossier ;
    n:=executeSql(st) ;
  end ;

  //dossier idem ajouté
  lstC := TListBox(getcontrol('LstCreate')) ;
  for i:= 0 to lstC.Items.count-1 do
  begin
    st := 'insert into DPDTPROCESS ' +
          '(  DDP_DPDTPROCESS, DDP_NODOSSIER, DDP_UTILISATEUR, DDP_METIER, ' +
          '   DDP_FINALITE, DDP_MISSION, DDP_SEUIL, DDP_DATEDU, DDP_DATEAU '+
          ') ' +
          'values (' +
          intTostr(getCompteurTable) + ', ' +
          '"' + lstC.Items[i] + '", ' +
          '"' + getControlText('DDP_UTILISATEUR') + '", ' +
          '"' + getControlText('DDP_METIER') + '", ' +
          '"' + getControlText('DDP_FINALITE') + '", ' +
          '"' + getControlText('DDP_MISSION') + '", ' +
          '"' + StrfPoint(strToFloat(getControlText('DDP_SEUIL'))) + '", ' +
          '"' + usdateTime(strToDate(getControlText('DDP_DATEDU'))) + '", ' +
          '"' + usdateTime(strToDate(getControlText('DDP_DATEAU'))) + '" ' +
          ')' ;
    n:=ExecuteSql(st) ;
  end ;
  ModifAnnexe:=false ;
end ;

Initialization
  registerclasses([TOM_DPDTPROCESS]) ;

end.
