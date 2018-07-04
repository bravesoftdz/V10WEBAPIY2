unit TofMulBudRub;

interface
uses  Controls,StdCtrls,ComCtrls,Graphics,Classes,sysutils,dbTables,
      HCtrls,HEnt1,HMsgBox,UTOF,LookUp,HDB,FE_Main,Vierge,HRichOLE,UtilDiv,RubFam ;
Type
     TOF_ConvRubBud = class (TOF)
     private
       IncOK: Boolean ;
       LaFamille: THValComboBox ;
       Pages: TPageControl ;
       FListe: THDBGrid ;
       LSelected,LRubrique: TStringList ;
       procedure OnRubriqueClick(Sender : TObject) ;
       procedure OnBOuvrirClick(Sender : TObject) ;
       procedure OnBZoomClick(Sender : TObject) ;
       procedure ChargeLRubrique ;
       procedure ChargeLSelected ;
       procedure IsIncremental ;
       procedure MiseAJourQuery(var Q2 : TQuery) ;
       function ExisteRub(Rub : string) : boolean ;
     public
       procedure OnArgument  (stArgument : String ) ; override ;
       procedure OnLoad ; override ;
       procedure OnClose ; override ;
     end;

implementation
procedure TOF_ConvRubBud.OnArgument (stArgument : String ) ;
var CEdit : THEdit ; BButton : TButton ;
begin
inherited ;
CEdit:=THedit(GetControl('RB_RUBRIQUE')) ;
if (CEdit<>nil) and not assigned(CEdit.OnElipsisClick) then CEdit.OnElipsisClick:=OnRubriqueClick ;
BButton:=TButton(GetControl('BOUVRIR')) ;
if (BButton<>nil) then BButton.OnClick:=OnBOuvrirClick ;
BButton:=TButton(GetControl('BZOOM')) ;
if (BButton<>nil) and not assigned(BButton.OnClick) then BButton.OnClick:=OnBZoomClick ;
LaFamille:=THValComboBox(GetControl('LAFAMILLE')) ;
Pages:=TPageControl(GetControl('PAGES')) ;
FListe:=THDBGrid(GetControl('FLISTE')) ;
LRubrique:=TStringList.Create ;
LSelected:=TStringList.Create ;
end;

procedure TOF_ConvRubBud.OnLoad ;
Begin
inherited ;
end;

procedure TOF_ConvRubBud.OnClose ;
Begin
inherited ;
LRubrique.Free ;
LSelected.Free ;
end;

procedure TOF_ConvRubBud.OnRubriqueClick(Sender : TObject) ;
begin
LookUpList(TControl(Sender),'Liste des rubriques','RUBRIQUE','RB_RUBRIQUE','RB_LIBELLE','RB_NATRUB="BUD"','RB_RUBRIQUE',True,0) ;
end ;

procedure TOF_ConvRubBud.OnBOuvrirClick(Sender : TObject) ;
var i,j,NumInc : integer ;Q,Q2 : TQuery ; Familles : string ;
Begin
ChargeLRubrique ;
ChargeLSelected ;
IsIncremental ;
if LaFamille.Value='' then
  begin
  HShowMessage('2;Conversion des ruptures en rubriques;Vous devez renseigner la famille de rubriques;W;O;O;O','','') ;
  Pages.ActivePage:=Pages.Pages[4] ; LaFamille.SetFocus ; Exit ;
  end ;
if IncOK then NumInc:=StrToInt(GetChoixCode('GRB','AUT')) else NumInc:=0 ;
for i:=0 to LSelected.Count-1 do
  begin
  if IncOk then while ExisteRub('GE:'+IntToStr(NumInc)) do Inc(NumInc) ;
  Q:=OpenSql('SELECT * FROM RUBRIQUE WHERE RB_RUBRIQUE="'+LSelected.Strings[i]+'"',False) ;
  if Not Q.Eof then
    begin
    Q2:=OpenSql('SELECT * FROM RUBRIQUE WHERE RB_RUBRIQUE="'+W_W+'"',False) ;
    Q2.Insert ;
    InitNew(Q2) ;
    for j:=0 to Q.fields.Count-1 do Q2.FieldValues[Q2.Fields[j].FieldName]:=Q.FieldValues[Q.Fields[j].FieldName] ;
    if IncOK then Q2.FindField('RB_RUBRIQUE').AsString:='GE:'+IntToStr(NumInc)
             else Q2.FindField('RB_RUBRIQUE').AsString:='GE:'+LSelected.Strings[i] ;
    Q2.FindField('RB_FAMILLES').AsString:=Familles ;
    MiseAJourQuery(Q2) ;
    Q2.Post ;
    ferme(Q2);
    if IncOk then Inc(NumInc) ;
    end ;
  ferme(Q) ;
  end ;
if IncOK then SetChoixCode('GRB','AUT',IntToStr(NumInc)) ;
Fliste.ClearSelected ;
end ;

procedure TOF_ConvRubBud.OnBZoomClick(Sender : TObject) ;
begin
FamillesRub ;
end ;

procedure TOF_ConvRubBud.ChargeLRubrique ;
var Q : TQuery ;
begin
if LRubrique<>nil then
  begin
  LRubrique.Clear ;
  Q:=OpenSql('SELECT RB_RUBRIQUE FROM RUBRIQUE ORDER BY RB_RUBRIQUE',True) ;
  while not Q.eof do  begin  LRubrique.Add(Q.Findfield('RB_RUBRIQUE').AsString) ; Q.Next ; end ;
  ferme(Q) ;
  end ;
end ;

procedure TOF_ConvRubBud.ChargeLSelected ;
var i : integer ; Q : TQuery ;
begin
if LSelected<>nil then
  begin
  LSelected.Clear ;
  Q:=TQuery(FListe.DataSource.DataSet) ;
  for i:=0 to FListe.nbSelected-1 do
    begin
    FListe.GotoLeBookmark(i) ;
    LSelected.Add(Q.Findfield('RB_RUBRIQUE').AsString) ;
    end ;
  end ;
end ;

procedure TOF_ConvRubBud.IsIncremental ;
var i : integer ;
begin
IncOk:=False ;
for i:=0 to LSelected.Count-1 do
  begin
  IncOK:=(length(LSelected.Strings[i])>14) ;
  if not IncOK then IncOK:=ExisteRub('GE:'+LSelected.Strings[i]) ;
  if IncOK then break ;
  end ;
end ;

procedure TOF_ConvRubBud.MiseAJourQuery(var Q2 : TQuery) ;
var CptTmp,ExcTmp : string ;
begin
Q2.FindField('RB_NATRUB').AsString:='CPT' ;
Q2.FindField('RB_BUDJAL').AsString:='' ;
// JLD + GP Pour XX_PREDEFINI
Q2.FindField('RB_PREDEFINI').AsString:='DOS' ; Q2.FindField('RB_NODOSSIER').AsString:='000000' ;
if V_PGI_Env<>Nil then if V_PGI_Env.ModeFonc<>'MONO' then Q2.FindField('RB_NODOSSIER').AsString:=V_PGI_ENV.NoDossier ;
if Q2.FindField('RB_TYPERUB').AsString='A/G' then
  begin
  Q2.FindField('RB_TYPERUB').AsString:='G/A' ;
  CptTmp:=Q2.FindField('RB_COMPTE1').AsString ;
  ExcTmp:=Q2.FindField('RB_EXCLUSION1').AsString ;
  Q2.FindField('RB_COMPTE1').AsString:=Q2.FindField('RB_COMPTE2').AsString ;
  Q2.FindField('RB_EXCLUSION1').AsString:=Q2.FindField('RB_EXCLUSION2').AsString ;
  Q2.FindField('RB_COMPTE2').AsString:=CptTmp ;
  Q2.FindField('RB_EXCLUSION2').AsString:=ExcTmp ;
  end ;
end ;

function TOF_ConvRubBud.ExisteRub(Rub : string) : boolean ;
var i : integer ;
begin
Result:=False ;
if LRubrique<>nil then
  begin
  for i:=0 to LRubrique.Count-1 do
    begin
    if Rub=LRubrique.Strings[i] then Result:=True ;
    if (Result=True) or (Rub<LRubrique.Strings[i]) then break ;
    end ;
  end ;
end ;

Initialization
registerclasses([TOF_ConvRubBud]);
end.
