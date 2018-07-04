//TOM_CPPROFILUSERC est utilisée pas les fiches PROFILUSER et PROFILUSER2
unit TomProfilUser;

interface
uses
    Classes, Dialogs, stdctrls, Sysutils, Math,
{$IFDEF EAGLCLIENT}
     MainEAgl,eFichList,
{$ELSE}
    db, dbctrls,{$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} FichList, HDB, FE_Main ,
{$ENDIF}
{$IFDEF MODENT1}
    CPTypeCons,
    CPProcMetier,
{$ENDIF MODENT1}
    UTOM, UTOB, HmsgBox, Hctrls, Ent1, HEnt1, EntGC   ,UentCommun,paramsoc,UtilRessource;

procedure YYLanceFiche_ProfilUser;

type
  TOM_CPPROFILUSERC = class(TOM)
    procedure OnUpdateRecord ; override ;
    procedure OnLoadRecord ; override ;
    procedure OnChangeField (F : TField); override ;
    procedure OnNewRecord  ; override ;
    procedure OnDeleteRecord ; override ;
    procedure OnArgument(Arguments : String ); override ;
    procedure OnClose ; override ;
    procedure CPU_TABLELIBRE1DblClick(Sender: TObject);
    procedure CPU_TABLELIBRE2DblClick(Sender: TObject);
    procedure CPU_USERChange(Sender: TObject);
{$IFNDEF EAGLCLIENT}
    procedure CAuxChange(Sender: TObject);
{$ENDIF}
  private
    TOBRESTBTP : TOB;
{$IFDEF EAGLCLIENT}
    TLGEN, TLAux, CGen,CAux,SaufGen,SaufAux : THEdit;
{$ELSE}
    TLGEN, TLAux, CGen,CAux,SaufGen,SaufAux : THDBEdit;
{$ENDIF}
    Responsable : THEdit;
    TypeUser: string;
    NivRestr : string; // JTR - Dépôt par défaut
    DepotLie : string;
    CpuGrp, CpuUser, Etab, Domaine, Depot: TCombobox;
    ForceEtab, ForceDomaine, PcpVte, PcpAch : TCheckBox;
    ForceDepot : TCheckBox; // JTR - Force le dépot
    function GetUserGroupe(User: string): string;
    {$IFNDEF EAGLCLIENT}
    function GetWhere(Cpt, Exclu, StTL : string; fb: TFichierBase ; Var Tous : Boolean): string;
    function CreerTob(sql: string): Tob;
    function CreerList(tobG, tobT: tob): TStringList;
    function IntersectionPresente(ListN, ListO: TStringList): boolean;
    function PresenceIntersection(CptG, ExcluG, CptT, ExcluT,StTLG,StTLT : string ; Var UserPB : String): boolean;
    Function AnalyseTL(St : String ; Var UserPB : String) : Boolean ;
    procedure SwapEnabled (St : String ; OkOk : Boolean) ;
    {$ENDIF !EAGLCLIENT}
    procedure ResponsableChange (Sender : TObject);
    procedure BRechResponsable(Sender: TObject);
End;

implementation

{$IFNDEF EAGLCLIENT}
 {$IFNDEF GCGC}
uses TabLiRub;
 {$ENDIF}
{$ENDIF}
var EstModifie : boolean;

procedure YYLanceFiche_ProfilUser;
begin
  AGLLanceFiche('BTP','BTPROFILTAB','','','ETA') ;
end;

procedure TransFormeCode(Var St : String ; PVirgule : Boolean);
begin
  if St = '' then Exit ;
  if PVirgule then
     While Pos(',', St) > 0 do
       St[Pos(',', St)] := ';'
  else
     While Pos(';', St) > 0 do
       St[Pos(';', St)] := ',';
  if PVirgule then
    St := St + ';'
  else
    Delete(St, Length(St), 1);
end;

{$IFNDEF GCGC}
Function ZoomTL(fb : tFichierBase ; Cpt : String) : String ;
Var CodDe,CodA,Stock,St : String ;
    i,j : Integer ;
begin
CodDe:='' ; CodA:='' ; Stock:='' ;
If Cpt<>'' Then
  BEGIN
  i:=Pos(':',Cpt) ;
  if i>0 then
     BEGIN
     CodDe:=Copy(Cpt,1,i-1) ;
     CodA:=Copy(Cpt,i+1,Length(CodDe)) ;
     Stock:=Copy(Cpt,2*Length(CodDe)+2,Length(Cpt)) ;
     TransformeCode(CodDe,True) ; TransformeCode(CodA,True) ;
     END else
     BEGIN
     i:=Pos('&',Cpt) ; j:=Pos('|',Cpt) ;
     if (i=0) And (j=0) then CodDe:=Cpt ;
     if (i<>0) And (j=0) then BEGIN CodDe:=Copy(Cpt,1,i-2) ; Stock:=Copy(Cpt,i,Length(Cpt)) ; END ;
     if (j<>0) And (i=0) then BEGIN CodDe:=Copy(Cpt,1,j-2) ; Stock:=Copy(Cpt,j,Length(Cpt)) ; END ;
     if (i<>0) And (j<>0)then BEGIN CodDe:=Copy(Cpt,1,Min(i,j)-2) ; Stock:=Copy(Cpt,Min(i,j),Length(Cpt)) ; END ;
     TransformeCode(CodDe,True) ; CodA:=CodDe ;
     END ;
  END ;
if Stock<>'' then
   if Stock[1]=',' then Stock:=Copy(Stock,2,Length(Stock)) ;
{$IFNDEF EAGLCLIENT}
ChoixTableLibrePourRub(fb,'',Stock,CodDe,CodA) ;
{$ENDIF}
TransformeCode(CodDe,False) ; TransformeCode(CodA,False) ;
if CodA<>CodDe then St:=CodDe+':'+CodA else St:=CodDe ;
if Stock<>'' then BEGIN if Stock[1]<>',' then Stock:=','+Stock ; St:=St+Stock ; END ;
Result:=St ;
end;
{$ENDIF}


procedure TOM_CPPROFILUSERC.CPU_TABLELIBRE1DblClick(Sender: TObject);
begin
  inherited;
{$IFNDEF GCGC}
  if DS.State = dsBrowse then DS.Edit;
  SetControlText('CPU_TABLELIBRE1', ZoomTL(fbGene, trim(TLGen.Text)));
{$ENDIF}
end;

procedure TOM_CPPROFILUSERC.CPU_TABLELIBRE2DblClick(Sender: TObject);
begin
  inherited;
{$IFNDEF GCGC}
  if DS.State = dsBrowse then DS.Edit;
  SetControlText('CPU_TABLELIBRE2', ZoomTL(fbAux, trim(TLAux.Text)));
{$ENDIF}
end;

function TOM_CPPROFILUSERC.GetUserGroupe(User: string): string;
var
  q: TQuery;
begin
  result := '';
  q := openSql('select US_GROUPE from UTILISAT where US_UTILISATEUR="' + User + '"', true,-1,'',true);
  if not q.eof then
    result := q.Fields[0].AsString;
  ferme(q);
end;

procedure TOM_CPPROFILUSERC.OnArgument(Arguments: String);
begin
  inherited;
  TOBRESTBTP := TOB.Create ('BPROFILUSERC',nil,-1);
  Ecran.HelpContext:=11000019;
  TypeUser := ReadTokenSt(Arguments);
  NivRestr := ReadTokenSt(Arguments); // JTR - Dépôt par défaut
  if NivRestr = '' then NivRestr := 'UTI'; // pour conserver le fonctionnement précédent
  DepotLie := ''; // JTR - Dépôts liés à l'éts si existe
{$IFDEF EAGLCLIENT}
  TLGen   := THEdit(GetControl('CPU_TABLELIBRE1'));
  TLAux   := THEdit(GetControl('CPU_TABLELIBRE2'));
  CGen    := THEdit(GetControl('CPU_COMPTE1'));
  CAux    := THEdit(GetControl('CPU_COMPTE2'));
  SaufGen := THEdit(GetControl('CPU_EXCLUSION1'));
  SaufAux := THEdit(GetControl('CPU_EXCLUSION2'));
{$ELSE}
  TLGen   := THDBEdit(GetControl('CPU_TABLELIBRE1'));
  TLAux   := THDBEdit(GetControl('CPU_TABLELIBRE2'));
  CGen    := THDBEdit(GetControl('CPU_COMPTE1'));
  CAux    := THDBEdit(GetControl('CPU_COMPTE2'));
  SaufGen := THDBEdit(GetControl('CPU_EXCLUSION1'));
  SaufAux := THDBEdit(GetControl('CPU_EXCLUSION2'));
  SetControlProperty('TCPU_TABLELIBRE2','WordWrap',FALSE) ;
  SetControlProperty('TCPU_TABLELIBRE1','WordWrap',FALSE) ;
  SetControlProperty('TCPU_EXCLUSION2','WordWrap',FALSE) ;
  SetControlProperty('TCPU_EXCLUSION1','WordWrap',FALSE) ;
  SetControlProperty('TCPU_COMPTE2','WordWrap',FALSE) ;
  SetControlProperty('TCPU_COMPTE1','WordWrap',FALSE) ;
  SetControlProperty('TCPU_TYPE','WordWrap',FALSE) ;
  if NivRestr = 'UTI' then
  begin
    SetControlVisible('TBPU_RESPONSABLE',true);
    SetControlVisible('BPU_RESPONSABLE',true);
    SetControlVisible('BPU_FORCERESP',true);
    Responsable := THEdit(getControl('BPU_RESPONSABLE'));
  end;
  if DS is TTable then // JTR - Dépôt par défaut
  begin
    if  NivRestr = 'GRP' then
      TTable(DS).Filter := 'CPU_TYPE='''+TypeUser+''' AND CPU_USER=''...'''
    else if NivRestr = 'UTI' then
      TTable(DS).Filter := 'CPU_TYPE='''+TypeUser+''' AND CPU_USER<>''...'''
    else
      TTable(DS).Filter := 'CPU_TYPE='''+TypeUser+'''';
    TTable(DS).Filtered:=TRUE;
  end ;
{$ENDIF}
  CpuGrp := TComboBox(GetControl('CPU_USERGRP'));
  CpuUser := TComboBox(GetControl('CPU_USER'));
  if CpuUser <> nil then
     CpuUser.OnChange := CPU_USERChange;
{$IFDEF EAGLCLIENT}
  TLGen := THEdit(GetControl('CPU_TABLELIBRE1'));
  TLAux := THEdit(GetControl('CPU_TABLELIBRE2'));
{$ELSE}
  TLGen := THDBEdit(GetControl('CPU_TABLELIBRE1'));
  TLAux := THDBEdit(GetControl('CPU_TABLELIBRE2'));
{$ENDIF EAGLCLIENT}
  if TLGen <> nil then
    TLGen.OnDblClick := CPU_TABLELIBRE1DblClick;
  if TLAux <> nil then
    TLAux.OnDblClick := CPU_TABLELIBRE2DblClick;
  Etab    := TComboBox(GetControl('CPU_ETABLISSEMENT'));
  ForceEtab    := TCheckBox(GetControl('CPU_FORCEETAB'));
  Domaine := TCombobox(GetControl('CPU_DOMAINE'));
  ForceDomaine := TCheckBox(GetControl('CPU_FORCEDOMAINE'));
  if V_PGI.NumVersionBase >= 631 then
  begin
    Depot  := TCombobox(GetControl('CPU_DEPOT'));
    ForceDepot := TCheckBox(GetControl('CPU_FORCEDEPOT'));
    PcpVte := TCheckBox(GetControl('CPU_PCPVENTE'));
    PcpAch     := TCheckBox(GetControl('CPU_PCPACHAT'));
    PcpAch.Hint := '';
    PcpVte.Hint := '';
  end;
  if CpuUser <> nil then CpuUser.OnChange := CPU_USERChange;
  if TLGen <> nil then TLGen.OnDblClick := CPU_TABLELIBRE1DblClick;
  if TLAux <> nil then TLAux.OnDblClick := CPU_TABLELIBRE2DblClick;
{$IFNDEF EAGLCLIENT}
  if CAux <> nil then CAux.OnChange := CAUXChange;
{$ENDIF EAGLCLIENT}

  // JTR - Dépôt par défaut - Active/Désactive champs si Groupes ou Utilisateurs
  CpuGrp.Visible := (NivRestr='GRP');
  ThLabel(GetControl('TCPU_USERGRP')).Visible := (NivRestr='GRP');
  CpuUser.Visible := (NivRestr='UTI');
  if V_PGI.NumVersionBase >= 631 then
  begin
    ThLabel(GetControl('TCPU_USER')).Visible := (NivRestr='UTI');
{$IFDEF NOMADESERVER}
    THLabel(GetControl('HTITREPCP')).Visible := (NivRestr='UTI');
    PcpVte.Visible := (NivRestr='UTI');
    PcpAch.Visible := (NivRestr='UTI');
{$ELSE}
 {$IFDEF NOMADE}
    THLabel(GetControl('HTITREPCP')).Visible := (NivRestr='UTI');
    PcpVte.Visible := (NivRestr='UTI');
    PcpAch.Visible := (NivRestr='UTI');
 {$ELSE NOMADE}
    THLabel(GetControl('HTITREPCP')).Visible := False;
    PcpVte.Visible := False;
    PcpAch.Visible := False;
 {$ENDIF NOMADE}
{$ENDIF NOMADESERVER}
  end;
  CpuUser.Visible := (NivRestr='UTI');
//  ForceEtab.Enabled := (NivRestr='UTI');
//  Domaine.Enabled := (NivRestr='UTI');
//  ForceDomaine.Enabled := (NivRestr='UTI');
  EstModifie := False;

 if (ctxAffaire in V_PGI.PGIContexte)  then
 begin
  if (ctxScot in V_PGI.PGIContexte)  then
  begin
    if GetControl('HTITREDOMAINE')    <> nil then   SetControlVisible('HTITREDOMAINE', False);
    if GetControl('CPU_DOMAINE')      <> nil then   SetControlVisible('CPU_DOMAINE', False);
    if GetControl('CPU_FORCEDOMAINE') <> nil then   SetControlVisible('CPU_FORCEDOMAINE', False);
  end;
  if GetControl('HTITREDEPOT')      <> nil then   SetControlVisible('HTITREDEPOT', False);
  if GetControl('CPU_DEPOT')        <> nil then   SetControlVisible('CPU_DEPOT', False);
  if GetControl('CPU_FORCEDEPOT')   <> nil then   SetControlVisible('CPU_FORCEDEPOT', False);
 end;

end;

procedure TOM_CPPROFILUSERC.OnChangeField(F: TField);
var Qry : TQuery;
begin
  inherited;
// PCS 12/05/2003 supprimé IFNDEF GCGC
  if ((NivRestr = 'UTI') or (NivRestr = '')) and (F.FieldName = 'CPU_USER') then
    SetField('CPU_USERGRP', GetUserGroupe(GetControlText('CPU_USER')));
  if (ctxMode in V_PGI.PGIContexte) and (TypeUser = 'ETA') and (F.FieldName = 'CPU_ETABLISSEMENT') then
  begin
    Qry := OpenSql('SELECT ET_DEPOTLIE FROM ETABLISS WHERE ET_ETABLISSEMENT = "'+GetField('CPU_ETABLISSEMENT')+'"',true,-1,'',true);
    if not Qry.Eof then
      DepotLie := Qry.Fields[0].AsString
      else
      DepotLie := '';
    Ferme(Qry);
  end;
end;

procedure TOM_CPPROFILUSERC.OnDeleteRecord;
begin
  inherited;
end;

procedure TOM_CPPROFILUSERC.OnLoadRecord;
var QQ : TQuery;
begin
  inherited;
  TOBRESTBTP.InitValeurs(false);
  if (TypeUser = 'ETA') and (GetField('CPU_TYPE')='') then
    begin
    ds.edit;
    SetField('CPU_TYPE', TypeUser);
    end;
  if NivRestr = 'GRP' then // JTR - eQualité 11790
    Ecran.caption := TraduireMemoire('Restrictions groupes utilisateurs')
    else
    Ecran.caption := TraduireMemoire('Restrictions utilisateurs');
  if NivRestr = 'UTI' then
  begin
    QQ := OpenSql ('SELECT * FROM BPROFILUSERC WHERE BPU_USER="'+GetField('CPU_USER')+'"',true,1,'',true);
    if not QQ.eof then TOBRESTBTP.SelectDB('',QQ);
    ferme (QQ);
    TOBRESTBTP.PutEcran(ecran);
    Responsable.OnChange := ResponsableChange;
    Responsable.OnElipsisClick := BRechResponsable;
  end;
  UpdateCaption(Ecran);
end;

procedure TOM_CPPROFILUSERC.OnNewRecord;
begin
  inherited;
  SetControlEnabled('CPU_USERGRP', false);
  SetControlEnabled('CPU_TYPE', false);
  SetField('CPU_TYPE', TypeUser);
  if ds.recordcount <= 0 then
     begin
     SetField('CPU_USER', V_PGI.USER);
     SetField('CPU_USERGRP', GetUserGroupe(V_PGI.USER));
     end;
  if TypeUser='ETA' then
     BEGIN
     SetField('CPU_USER','...') ;
     SetControlEnabled('CPU_USERGRP',True);
     {JP 27/11/2003 : Cette ligne exitait dans VDEV, ne sachant pas s'il faut la porter ...
                      je la porte ... en commentaires !!!
     SetField('CPU_USERGRP', GetUserGroupe(V_PGI.USER));}
     END ;
end;

procedure TOM_CPPROFILUSERC.OnUpdateRecord;
{$IFNDEF EAGLCLIENT} //JTR - Enlève conseils et avertissements
Var UserPB : String ;
{$ENDIF EAGLCLIENT}
begin
  inherited;
  if (TypeUser = 'ETA') then
  begin
    // JTR - Dépôt par défaut
    if NivRestr = 'GRP' then
    begin
      SetField('CPU_USER','...'); // JTR - Fonction en CWAS
      if GetControlText('CPU_USERGRP') = '' then
      begin
        PGIBox('Vous devez renseigner un groupe', 'Profils groupes utilisateurs');
        SetFocusControl('CPU_USERGRP');
        exit;
      end;
    end else
    if (NivRestr = 'UTI') and (GetControlText('CPU_USER') = '') then
    begin
      PGIBox('Vous devez renseigner un utilisateur', 'Profils utilisateurs');
      SetFocusControl('CPU_USER');
        exit;
    end;
    if V_PGI.NumVersionBase >= 631 then
    begin
      if (ctxMode in V_PGI.PGIContexte) and (GetControlText('CPU_DEPOT') <> '') and
           (pos(GetControlText('CPU_DEPOT'),DepotLie) = 0) and (DepotLie <>'')then
      begin
        PGIBox('Ce dépôt n''est pas lié à l''établissement.','Profils utilisateurs');
        SetControlText('CPU_DEPOT','');
        SetFocusControl('CPU_DEPOT');
          exit;
      end;
      if ForceDepot.Checked and (Depot.ItemIndex <= 0) then
      begin
        PGIBox('Vous devez renseigner un dépôt', 'Profils utilisateurs');
        SetFocusControl('CPU_DEPOT');
        exit;
      end;
    end;
    // Fin JTR
    if ForceEtab.Checked and (Etab.ItemIndex <= 0) then
    begin
      PGIBox('Vous devez renseigner un établissement', 'Profils utilisateurs');
      SetFocusControl('CPU_ETABLISSEMENT');
      Abort;
    end;
    EstModifie := True;
    
    if (NivRestr = 'UTI') then
    begin
      TOBRESTBTP.GetEcran(Ecran);
      if TOBRESTBTP.getString('BPU_USER')='' then TOBRESTBTP.SetString('BPU_USER',GetField('CPU_USER'));
      TOBRESTBTP.InsertOrUpdateDB(false);
    end;

    exit;
  end;
{$IFNDEF EAGLCLIENT}
  if GetControlText('CPU_TYPE') = '' then
  begin
    PGIBox('Vous devez renseigner un type de traitement', 'Profils utilisateurs');
    SetFocusControl('CPU_TYPE');
    Abort;
  end;
  if DS.State = dsinsert then
  begin
    if PresenceComplexe('CPPROFILUSERC',
                        ['CPU_USERGRP','CPU_USER','CPU_TYPE'],
                        ['=','=','='], [GetControlText('CPU_USERGRP'),
                        GetControlText('CPU_USER'),GetControlText('CPU_TYPE')],
                        ['S','S','S']) then
    begin
      PGIBox('Vous ne pouvez pas créer ce profil. Il existe déjà.', 'Profils utilisateurs');
      Abort;
    end;
  end;
  if PresenceIntersection(CGen.Text, SaufGen.Text,
                          CAux.Text, SaufAux.Text,
                          TLGen.Text, TLAux.Text,USerPB) then
  begin
     PGIBox('Il y a des intersections dans les fourchettes des comptes ('+UserPb+')', 'Profils utilisateurs');
     Abort;
  end;
  If (TypeUser='CLI') Or (TypeUser='FOU') Then
  BEGIN
    If (TLAux.Text='') And (CAux.Text='') Then
    BEGIN
      PGIBox('Vous devez renseigner au moins une zone "auxiliaire"', 'Profils utilisateurs');
      SetFocusControl('CPU_COMPTE2');
      Abort;
    END ;
  END ;
{$ENDIF}
end;

procedure TOM_CPPROFILUSERC.OnClose ;
begin
inherited ;
ChargeProfilUser ;
TOBRESTBTP.free;
if EstModifie  then
  PGIBox('Vos modifications seront actives à la prochaine connexion','Profils utilisateurs');
end ;

{$IFNDEF EAGLCLIENT}
function TOM_CPPROFILUSERC.GetWhere(Cpt, Exclu, StTL: string; fb: TFichierBase ; Var Tous : Boolean): string;
var WhereCpt,Wheresauf, WhereTL : string;
    OkWhere : Boolean ;
begin
  Result:='' ; Tous:=FALSE ;
  WhereCpt:='' ; Wheresauf:='' ;  WhereTL:='' ; OkWhere:=FALSE ;
  if Cpt<>'' then
    BEGIN
    OkWhere:=TRUE ;
    If Trim(Cpt)='TOUS' Then
      BEGIN
      Tous:=TRUE ;
      Case fb Of
        fbGene : WhereCpt:='(G_GENERAL<>"'+W_W+'") AND ' ;
        fbAux : WhereCpt:='(T_AUXILIAIRE<>"'+W_W+'") AND ' ;
        END ;
      END Else WhereCpt:=' ('+AnalyseCompte(Cpt, fb, false, false, true)+') AND ';
    END ;
  if Exclu<>'' then BEGIN OkWhere:=TRUE ; WhereSauf:=' ('+AnalyseCompte(Exclu, fb, true, false, true)+') AND '; END ;
  If StTL<>'' Then BEGIN OkWhere:=TRUE ; WhereTL:=' ('+AnalyseCompte(StTL,fb,False,TRUE,FALSE)+') AND ' ; END ;
  if OkWhere then BEGIN Result:=WhereCpt+WhereSauf+WhereTL ; Delete(Result,Length(Result)-4,4) ;
  END ;
end;

function TOM_CPPROFILUSERC.CreerTob(sql: string): Tob;
var
  q: TQuery;
begin
  q := OpenSql(sql, true,-1,'',true);
  result := Tob.Create('TOBCPT', nil, -1);
  try
    while not q.eof do begin
      TOB.CreateDB('CPT', result, -1, q);
      q.next ;
    end;
  finally
    ferme(q);
  end;
end;

function TOM_CPPROFILUSERC.CreerList(tobG, tobT: tob): TStringList;
var i, j: integer;
begin
result := TStringList.Create;
if (tobG.Detail.count <= 0) and (tobT.Detail.count <= 0) then exit;
if (tobG.Detail.count > 0) and (tobT.Detail.count <= 0) then
  begin
  for i := 0 to tobG.Detail.count - 1 do result.Add(tobG.Detail[i].GetValue('G_GENERAL'));
  exit;
  end;
if (tobG.Detail.count <= 0) and (tobT.Detail.count > 0) then
  begin
  for i := 0 to tobT.Detail.count - 1 do result.Add(tobT.Detail[i].GetValue('T_AUXILIAIRE')); exit;
  end;
for i := 0 to tobG.Detail.count - 1 do
  for j := 0 to tobT.Detail.Count - 1 do
    begin
    If tobG.Detail[i].GetValue('G_COLLECTIF')='X'
      Then result.Add(tobG.Detail[i].GetValue('G_GENERAL') + ',' + tobT.Detail[j].GetValue('T_AUXILIAIRE'))
      Else result.Add(tobT.Detail[j].GetValue('T_AUXILIAIRE')) ;
    end;
result.sorted := true;
end;
{$ENDIF !EAGLCLIENT}

{   Construire les couples: CN = Generaux X Auxiliaires pour une nouvelle enregistement
    Construire les couples: CO = Generaux X Auxiliaires pour chaque enregistrement dans le fichier
    Si (un couple de CN) appartient à CO, alors INTERSECTION}


procedure TOM_CPPROFILUSERC.CPU_USERChange(Sender: TObject);
begin
{$IFNDEF EAGLCLIENT}
OnChangeField(ds.FieldByName('CPU_USER'));
{$ENDIF}
end;

{$IFDEF EAGLCLIENT}
// AFAIREEAGL
(*
function TOM_CPPROFILUSERC.PresenceIntersection(CptG, ExcluG, CptT, ExcluT,StTLG,StTLT : string; Var UserPB : String): boolean;
BEGIN
Result:=False ;
END ;

function TOM_CPPROFILUSERC.IntersectionPresente(ListN, ListO: TStringList): boolean;
BEGIN
Result:=False ;
END ;

Function TOM_CPPROFILUSERC.AnalyseTL(St : String; Var UserPB : String) : Boolean ;
BEGIN
Result:=False ;
END ;

procedure TOM_CPPROFILUSERC.SwapEnabled (St : String ; OkOk : Boolean) ;
BEGIN
END ;
*)
{$ELSE}
function TOM_CPPROFILUSERC.IntersectionPresente(ListN, ListO: TStringList): boolean;
var
   i: integer;
begin
  result := false;
  for i := 0 to ListN.Count - 1 do
    if ListO.IndexOf(ListN[i]) >= 0 then begin
      result := true;
      break;
    end;
end;

Function TOM_CPPROFILUSERC.AnalyseTL(St : String; Var UserPB : String) : Boolean ;
Var SQL : String ;
    TobTL : TOB ;
    i : Integer ;
    StLu : String ;
//****,01F.,0000,----,#,#,#,---,#,#:****,01N.,0000,----,#,#,#,---,#,#
BEGIN
Result:=FALSe ; UserPB:='' ;
sql:='select CPU_TABLELIBRE2,CPU_USER from CPPROFILUSERC where CPU_USER<>"'+GetControlText('CPU_USER')+ '"'
    +' And CPU_TYPE="'+ GetControlText('CPU_TYPE')    + '"';
TobTL := CreerTob(sql);
for i:=0 to tobTL.Detail.count - 1 do
  BEGIN
  StLu:=TobTL.Detail[i].GetValue('CPU_TABLELIBRE2') ;
  If CompareTL(St,StLu) Then BEGIN Result:=TRUE ; UserPB:=TobTL.Detail[i].GetValue('CPU_USER') ; Break ; END ;
  END ;
TobTL.Free ;
END ;

function TOM_CPPROFILUSERC.PresenceIntersection(CptG, ExcluG, CptT, ExcluT,StTLG,StTLT : string; Var UserPB : String): boolean;
var
   where, sql: string;
   tobG, tobT: Tob;
   lstN, lstO: TStringList;
   q: TQuery;
   CptGLu,CptTLu,ExcluGLu,ExcluTLu : String ;
   Tous : Boolean ;
   TousLesAux,TousLesAuxLu : Boolean ;
begin
  UserPB:='' ;
  result := false;
  If AnalyseTL(StTLT,UserPB) Then BEGIN Result:=TRUE ; Exit ; END ;
  where := GetWhere(CptG, ExcluG, StTLG, fbGene,Tous);
  sql := 'select G_GENERAL,G_COLLECTIF from GENERAUX';
  if where='' then Where:=' G_GENERAL="'+W_W+'" ' ;
  sql := sql + ' where ' + where;
  TobG := CreerTob(sql);
  where := GetWhere(CptT, ExcluT, StTLT, fbAux,Tous); TousLesAux:=Tous ;
  sql := 'select T_AUXILIAIRE from TIERS';
  if where='' then Where:=' T_AUXILIAIRE="'+W_W+'" ' ;
  sql := sql + ' where ' + where;
  TobT := CreerTob(sql);
  lstN := CreerList(tobG, tobT);
  tobG.Free;
  tobT.Free;
  sql := 'select * from CPPROFILUSERC where CPU_USER<>"'    + GetControlText('CPU_USER')    + '"'
                                     + ' And CPU_TYPE="'    + GetControlText('CPU_TYPE')    + '"';

  q := OpenSql(sql, true,-1,'',true);
  lstO := nil;
  try
  while not q.eof do
    begin
    If TousLesAux Then BEGIN result := true; UserPB:=q.fieldbyname('CPU_USER').asstring ; break; end;
    CptGLu:=q.fieldbyname('CPU_COMPTE1').asstring ; ExcluGLu:=q.fieldbyname('CPU_EXCLUSION1').asstring ;
    CptTLu:=q.fieldbyname('CPU_COMPTE2').asstring ; ExcluTLu:=q.fieldbyname('CPU_EXCLUSION2').asstring ;
    where := GetWhere(CptGLu,ExcluGLu, q.fieldbyname('CPU_TABLELIBRE1').asstring, fbGene,Tous);
    sql := 'select G_GENERAL,G_COLLECTIF from GENERAUX';
    if where='' then Where:=' G_GENERAL="'+W_W+'" ' ;
    sql := sql + ' where ' + where;
    TobG := CreerTob(sql);
    where := GetWhere(CptTLu, ExcluTLu, q.fieldbyname('CPU_TABLELIBRE2').asstring, fbAux,Tous); TousLesAuxLu:=Tous ;
    If TousLesAuxLu Then BEGIN result := true; UserPB:=q.fieldbyname('CPU_USER').asstring ; break; end;
    sql := 'select T_AUXILIAIRE from TIERS';
    if where='' then Where:=' T_AUXILIAIRE="'+W_W+'" ' ;
    sql := sql + ' where ' + where;
    TobT := CreerTob(sql);
    lstO := CreerList(tobG, tobT);
    tobG.Free;
    tobT.Free;
    if IntersectionPresente(lstN, lstO) then begin result := true; UserPB:=q.fieldbyname('CPU_USER').asstring ; break; end;
    q.next;
    end;
  finally
    ferme(q);
    lstO.Free;
    lstN.Free;
  end;
end;

procedure TOM_CPPROFILUSERC.SwapEnabled (St : String ; OkOk : Boolean) ;
BEGIN
If OkOk then
  BEGIN
  SetControlEnabled(St,TRUE) ; SetControlEnabled('T'+St,TRUE) ;
  END Else
  BEGIN
  SetControlEnabled(St,FALSE) ; SetControlEnabled('T'+St,FALSE) ;
  If GetField(St)<>'' Then
    BEGIN
    if DS.State = dsBrowse then DS.Edit;
    SetField(St,'') ;
    END ;
  END ;
END ;

procedure TOM_CPPROFILUSERC.CAUXChange(Sender: TObject);
begin
SwapEnabled('CPU_TABLELIBRE2',TRUE) ; SwapEnabled('CPU_TABLELIBRE1',TRUE) ;
SwapEnabled('CPU_EXCLUSION2',TRUE) ;
If Trim(THDBEDit(Sender).Text)='TOUS' Then
  BEGIN
  SwapEnabled('CPU_TABLELIBRE2',FALSE) ; SwapEnabled('CPU_TABLELIBRE1',FALSE) ;
  SwapEnabled('CPU_EXCLUSION2',FALSE) ;
  END ;
end;
{$ENDIF}


procedure TOM_CPPROFILUSERC.ResponsableChange(Sender: TObject);
begin
  if not (DS.State in [dsInsert, dsEdit]) then DS.edit;
end;

procedure TOM_CPPROFILUSERC.BRechResponsable(Sender: TObject);
Var QQ  : TQuery;
    SS  : THCritMaskEdit;
begin

  if GetParamSocSecur('SO_AFRECHRESAV', True) then
  begin
    SS := THCritMaskEdit.Create(nil);
    GetRessourceRecherche(SS,'ARS_RESSOURCE=' + Responsable.text + ';TYPERESSOURCE=SAL', '', '');
    if (SS.Text <> Responsable.text) then
    begin
      if SS.text = '' then ss.text := Responsable.text;
      if not (DS.State in [dsInsert, dsEdit]) then DS.edit;
    end;
    Responsable.text  := SS.Text;
    SS.Free;
  end
  else
    GetRessourceRecherche(THCritMaskEdit(Responsable),'ARS_TYPERESSOURCE="SAL"', '', '');

end;


Initialization
registerclasses([TOM_CPPROFILUSERC]);


end.

