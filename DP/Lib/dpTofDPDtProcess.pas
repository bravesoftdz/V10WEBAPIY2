{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 28/07/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : DPDTPROCESS ()
Mots clefs ... : TOF;DPDTPROCESS
*****************************************************************}
(*
A faire :
  openSql en eAgl
*)

Unit dpTofDPDtProcess ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     Fe_Main,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     MaineAgl,
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
     stat ;

function CanDoProcess(Dossier, User, Finalite, Mission : string) : boolean ;

Type
  TOF_DPDTPROCESS = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    procKeyDownDecla : TKeyEvent ;
    //MesChamps : array of string ;
    procedure OnAfterShow;
    function  getCurrentId:string;
    function  LitTV(Fld : string; row : integer; var Val : variant):boolean;

    procedure DDP_Metier_OnClick (Sender: TObject);
    procedure TV_OnDblClick (Sender: TObject);
    procedure bDelete_OnClick (sender : TObject) ;
    procedure bInsert_OnClick (sender : TObject) ;
    procedure bDupliquer_OnClick (sender : TObject) ;
    procedure bReaffecter_OnClick (sender : TObject) ;
    procedure Form_OnKeyDown (Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure razCriteres ;
    procedure setLocalEvent ;
    //procedure initMesChamps (sql : string) ;
    procedure initColVisible ;
    procedure DeleteDt (cle:string);
    procedure setPos (st:string) ;
  end ;


Implementation

uses
  UtobView, hQry, uSatUtil, math, htb97, windows, LookUp, choisir ;

procedure TOF_DPDTPROCESS.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_DPDTPROCESS.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_DPDTPROCESS.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_DPDTPROCESS.OnLoad ;
begin
  Inherited ;
  SetControlText('LblFinalite', '') ;
  SetControlText('LblUser', '') ;
  SetControlText('LblNoDossier', '') ;
end ;

function tm(s:string) : string ;
begin
  result:='"' + FirstMajuscule(TraduireMemoire(s)) + '"' ;
end ;

procedure TOF_DPDTPROCESS.OnArgument (S : String ) ;
var sql : string ;
begin
  razCriteres ;
  Inherited ;
  sql := 'select (DDP_NODOSSIER + " - " + DOS_LIBELLE) as ' + tm('Dossier') + ', ' +
         'DDP_UTILISATEUR + " - " + US_LIBELLE ' + tm ('Utilisateur') + ', ' +
         'MET.CO_LIBELLE ' + tm ('Métier') + ',  MIS.CO_LIBELLE ' + tm ('Mission') + ',' +
         'DPC.CO_LIBELLE '+ tm('Objet') + ', PR.DDP_SEUIL ' + tm ('Seuil') + ', ' +
         'DDP_DPDTPROCESS '#13 +
         'from DPDTPROCESS PR '#13 +
         'inner join UTILISAT US on PR.DDP_UTILISATEUR = US.US_UTILISATEUR '#13 +
         'inner join DOSSIER D on D.DOS_NODOSSIER = PR.DDP_NODOSSIER '#13 +
         'inner join COMMUN MIS on MIS.CO_TYPE="MIS" and PR.DDP_MISSION = MIS.CO_CODE '#13+
         'inner join COMMUN DPC on DPC.CO_TYPE="DPC" and PR.DDP_FINALITE = DPC.CO_CODE '#13+
         'inner join COMMUN MET on MET.CO_TYPE="MER" and PR.DDP_METIER = MET.CO_CODE ';
         ; //a voir : filtrage comme DOSSIERGRP ?
  //InitMesChamps (sql);
  setControlText('FSQL', sql) ;
  setLocalEvent ;
  TFStat(Ecran).OnAfterFormShow := OnAfterShow;

end ;

{procedure TOF_DPDTPROCESS.initMesChamps (sql : string) ;
var q:tQuery;
    i : integer ;
begin
  sql := sql + ' where 1=2' ;
  q:=opensql(sql, true) ;
  SetLength(MesChamps, q.FieldCount) ;
  for i:=0 to q.FieldCount-1 do
    MesChamps[i] := q.Fields[i].FieldName ;
  Ferme(q) ;
end ;}


procedure TOF_DPDTPROCESS.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_DPDTPROCESS.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_DPDTPROCESS.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_DPDTPROCESS.OnAfterShow;
begin
  TFStat(Ecran).ForcePresentationChange := True ;
  ChangeListeCrit(TFStat(Ecran), false);
  initColVisible ;
end;

procedure TOF_DPDTPROCESS.DDP_Metier_OnClick(Sender: TObject);
var finalite : thEdit;
    Metier : thValComboBox ;
begin

  finalite := thEdit(getControl('DDP_FINALITE')) ;
  Metier := thValComboBox(getControl('DDP_METIER')) ;

  if (finalite=nil) or (Metier=nil) then exit ;

  finalite.plus := ' and CO_LIBRE = "' + getControlText('DDP_METIER') + '" ' ;
end ;

procedure TOF_DPDTPROCESS.setLocalEvent ;
begin
  procKeyDownDecla := TFStat(Ecran).OnKeydown ;
  TFStat(Ecran).OnKeydown := Form_OnKeyDown ;

  THEdit(GetControl('DDP_METIER')).OnClick := DDP_Metier_OnClick;
  TTobViewer(GetControl('TV')).OnDblClick := TV_OnDblClick;
  TButton(GetControl('BDELETE')).OnClick := bDelete_OnClick;
  TButton(GetControl('BINSERT')).OnClick := bInsert_OnClick;
  TButton(GetControl('BREAFFECTER')).OnClick := bReaffecter_OnClick;
  TButton(GetControl('BDUPLIQUER')).OnClick := bDupliquer_OnClick;

end ;

procedure TOF_DPDTPROCESS.RazCriteres ;
begin
  SetControlProperty('DDP_UTILISATEUR', 'Text','') ;
  SetControlProperty('DDP_NODOSSIER', 'Text','') ;
  SetControlProperty('DDP_FINALITE', 'Text','') ;
  SetControlProperty('DDP_METIER', 'Value','') ;
  SetControlProperty('DDP_MISSION', 'Value','') ;
end ;

procedure TOF_DPDTPROCESS.initColVisible ;
var i : integer ;
    tbv : TTobViewer ;
begin
  tbv := TTobViewer(GetControl('TV')) ;

  {
  for i :=  Low(MesChamps) to High(MesChamps) do
    if MesChamps[i][1] = NoView then
      tbv.ColWidths[tbv.ColIndex(MesChamps[i])] := -1 ;
  }

end ;

function TOF_DPDTPROCESS.getCurrentId : string;
var Vw : TTobViewer ;
    c : integer ;
begin
  Vw := TTobViewer(GetControl('TV')) ;
  c:= Vw.ColIndex('DDP_DPDTPROCESS') ;
  if c>-1 then result := Vw.GetValue(c, Vw.CurrentRow)
          else pgiInfo ('La colonne Index doit être incluse dans la présentation.');
end ;

procedure TOF_DPDTPROCESS.TV_OnDblClick (Sender: TObject) ;
var res, st : string ;
begin
  st := getCurrentId ;
  res := AglLanceFiche('DP', 'DPDTPROCESS_FIC', 'DDP_DPDTPROCESS=' + st, st, 'ACTION=MODIFICATION') ;
  if pos ('REFRESH;', res) > 0 then
  begin
    TButton(GetControl('BCHERCHE')).OnClick (nil) ;
    setPos(res) ;
  end ;
end ;

procedure TOF_DPDTPROCESS.bDelete_OnClick (sender : TObject) ;
var r:integer ;
begin
  if not pgiAsk('Confirmez-vous la suppression ?') = mrYes then exit ;
  r:=TTobViewer(GetControl('TV')).CurrentRow ;
  DeleteDt('DDP_DPDTPROCESS='+ getCurrentId) ;
  if r>TTobViewer(GetControl('TV')).RowCount then dec(r) ;
  TTobViewer(GetControl('TV')).SetRowSelect(r) ;
end ;

procedure TOF_DPDTPROCESS.bInsert_OnClick (sender : TObject) ;
var res : string ;
begin
  res := AglLanceFiche('DP', 'DPDTPROCESS_FIC', '', '', 'ACTION=CREATION') ;
  if pos ('REFRESH;', res) > 0 then TButton(GetControl('BCHERCHE')).OnClick (nil) ;
  setPos(res) ;
end ;

procedure TOF_DPDTPROCESS.bDupliquer_OnClick (sender : TObject) ;
var res : string ;
begin
  res := AglLanceFiche('DP', 'DPDTPROCESS_FIC', '', '', 'ACTION=CREATION;DUPLI;'+GetCurrentId) ;
  if pos ('REFRESH;', res) > 0 then TButton(GetControl('BCHERCHE')).OnClick (nil) ;
  setPos(res) ;
end ;

function  TOF_DPDTPROCESS.LitTV(Fld : string; row : integer; var Val : variant):boolean;
var Vw : TTobViewer ;
    col : integer ;
begin
  result:=false ;
  try
    Vw := TTobViewer(GetControl('TV')) ;
    col:= Vw.ColIndex(Fld) ;
    if col>-1 then
    begin
      Val := Vw.GetValue(col, row);
      result:=true ;
    end ;
  except
  end ;
  if not result then
    PgiInfo('Le champ ' + Fld + ' est absent de la vue. '#13 +
            'Il faut l''ajouter dans la représentation pour effectuer le traitement.');
end ;


procedure TOF_DPDTPROCESS.bReaffecter_OnClick (sender : TObject) ;
var UserIni : variant ;
    CodeUserIni, CodeUserDest : string ;
    n : integer ;
    sl : HTStringList ;
    Selection:TArrayofBoolean ;
begin

  if not LitTV('Utilisateur',TTobViewer(GetControl('TV')).CurrentRow, UserIni) then exit ;
  CodeUserIni := gtfs(UserIni, ' - ', 1);

  if pgiask ('Vous souhaitez changer l''utilisateur ' + UserIni + ' par l''utilisateur '#13 +
             'que vous allez sélectionner. Confirmez-vous le traitement?') <> mrYes then exit ;

  sl := HTStringList.create ;
  Tablette2TStringList('TTUTILISATEUR', 'US_UTILISATEUR<>"'+CodeUserIni+'"', sl) ;
  n := ChoisirDansLaListe('Changement de l''utilisateur ' + UserIni + ' par :', TStrings(sl), 0, false , Selection) ;
  if n=-1 then exit ;
  CodeUserDest:=gtfs(sl[n],#9, 1) ;
  sl.free ;
  RemplirListe('TTUTILISATEUR','');//Retabli la liste des users en mémoire

  if CodeUserDest='' then exit ;
  n := ExecuteSql('update DPDTPROCESS set DDP_UTILISATEUR="' + CodeUserDest + '" where DDP_UTILISATEUR="' + CodeUserIni + '"') ;
  PgiBox (intTostr(n) + ' enregistrement(s) modifié(s).') ;
  if n>0 then TButton(GetControl('BCHERCHE')).OnClick (nil) ;
end ;

procedure TOF_DPDTPROCESS.DeleteDt (cle:string);
var st, sWhere : string ;
    i:integer ;
begin //cle : token champ=valeur; au format du param "cle" d'aglLanceFiche
  st := '*'; i:=1 ;
  while st <> '' do
  begin
    st := gtfs (cle,';',i);
    if st='' then break ;
    if sWhere <> '' then sWhere := sWhere + ' and ' ;
    sWhere := sWhere + gtfs (st,'=',1) + ' = "' + gtfs (st,'=',2) + '" ' ;
    inc (i) ;
  end ;
  if executeSql('delete DPDTPROCESS where ' + sWhere) >=1 then TButton(GetControl('BCHERCHE')).OnClick (nil) ;
end ;

procedure TOF_DPDTPROCESS.setPos (st:string) ;
var TbC : array [0..0] of string ;
    TbV : array [0..0] of variant ;
    n : integer ;
begin
  if pos('DDP_DPDTPROCESS;', st) = 0 then exit ;
  n:= postfs(st, ';', 'DDP_DPDTPROCESS') ;
  TbV[0] := gtfs(st, ';', n+1) ;
  TbC[0] := 'DDP_DPDTPROCESS' ;
  n := TTobViewer(GetControl('TV')).FindRowWithKey(TbC, TbV) ;
  TTobViewer(GetControl('TV')).SetRowSelect(n) ;
end ;

function CanDoProcess(Dossier, User, Finalite, Mission : string) : boolean ;
//finalite cf Tablette DPFINALITE, Mission cf Tablette DPMISSION
var q: TQuery ;
    n : integer ;
begin

  result:=true;

  //si  aucun droit défini sur le dossier alors on considère le droit Ok (Judicieux?)
  {
  q:=openSql('select count(*) from DPDTPROCESS where DDP_NODOSSIER="' + Dossier + '"', true);
  n := q.fields[0].asInteger ;
  ferme(q) ;
  if n=0 then exit ;
  }

  q:=openSql('select count(*) from DPDTPROCESS '+
             'where DDP_NODOSSIER="'+ Dossier +'" and DDP_UTILISATEUR="'+ User +'" ' +
             'and DDP_FINALITE="'+ Finalite +'" and DDP_MISSION="'+ Mission +'" ' +
             'and DDP_DATEDU >= "'+ usDateTime(Now) +'" and DDP_DATEAU <= "' + usDateTime(Now) + '"'
             , true);
  n := q.fields[0].asInteger ;
  ferme(q) ;
  if n=0 then result := false ;

end ;

procedure TOF_DPDTPROCESS.Form_OnKeyDown (Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (key = vk_delete) and (ssCtrl in Shift) and GetControl('bDelete').visible then
  begin
    key := 0;
    TToolbarButton97(GetControl('bDelete')).OnClick (sender) ;
  end
  else if (key = vk_Imprime) and (ssCtrl in Shift) then
  begin
    key := 0;
    TToolbarButton97(GetControl('bImprimer')).OnClick (sender) ;
  end
  else if (key = vk_Nouveau) and (ssCtrl in Shift)
   and GetControl('bInsert').visible and GetControl('bInsert').Enabled then
  begin
    key := 0;
    TToolbarButton97(GetControl('bInsert')).OnClick (sender) ;
  end
  else
    procKeyDownDecla (Sender, Key,  Shift) ;
end ;

Initialization
  registerclasses ( [ TOF_DPDTPROCESS ] ) ;
end.

