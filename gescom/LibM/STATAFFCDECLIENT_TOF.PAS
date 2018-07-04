{***********UNITE*************************************************
Auteur  ...... : A. CATHELINEAU
Créé le ...... : 21/11/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : STATAFFCDECLIENT ()
Mots clefs ... : TOF;STATAFFCDECLIENT
*****************************************************************}
Unit STATAFFCDECLIENT_TOF ;

Interface

Uses StdCtrls,Controls,Classes,stat,Windows,
{$IFNDEF EAGLCLIENT}
     db,dbtables,DBGrids,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,UTobView,HQry,UTOB,
     Fe_Main,UtilArticle,UtilGC,EntGC,AGLInit,UtilDimArticle ;

Type
  TOF_STATAFFCDECLIENT = Class (TOF)
   private
        TV1 : TTobViewer;
        FTV : TFStat ;
        TOBDim,TOBAffect : TOB ;
        TypeAction : String ;

    procedure TVOnDblClick (Sender: TObject ) ;
    procedure MiseAJourTable ;
    procedure TraiterArticleDimensionne ;
    procedure TraiterArticleUnique ;
    procedure RemplirTOBDim ( CodeArticle : String; Commande : integer ) ;
    procedure TOBDimToTV(Commande: Integer) ;
    //procedure TVOnLineEnter (Sender: TObject ) ;

   public
    procedure OnArgument (S : String ); override ;
    procedure OnLoad                  ; override ;
    procedure OnUpdate                ; override ;
    procedure OnClose                 ; override ;
  end ;


Implementation

procedure TOF_STATAFFCDECLIENT.OnLoad ;
begin
  Inherited ;
TOBDim:=TOB.Create('',Nil,-1) ;
end ;

procedure TOF_STATAFFCDECLIENT.OnArgument (S : String ) ;
var iCol , Nbr : integer ;
    stIndice, stChampsCompl : string ;
begin
  Inherited ;
  TypeAction := S ;
  TV1 := TTobViewer(getcontrol('TV'));
  TV1.OnDblClick := TVOnDblClick ;
  FTV := TFStat(Ecran) ;

   // Paramétrage des libellés des familles, stat. article et de la collection
  ChangeLibre2 ('TGA_COLLECTION', Ecran) ;
  for iCol := 1 to 3 do
  begin
    stIndice := IntToStr (iCol) ;
    ChangeLibre2 ('TGA_FAMILLENIV' + stIndice, Ecran) ;
  end ;
  stChampsCompl := '' ;
  if (ctxMode in V_PGI.PGIContexte) and (GetPresentation = ART_ORLI) then
  begin
    for iCol := 4 to 8 do
    begin
      stIndice := IntToStr (iCol) ;
      if ChangeLibre2 ('TGA2_FAMILLENIV' + stIndice, Ecran) then stChampsCompl := stChampsCompl + ',GA2_FAMILLENIV' + stIndice ;
    end;
    for iCol:=1 to 2 do
    begin
      stIndice := IntToStr (iCol) ;
      if ChangeLibre2 ('TGA2_STATART' + stIndice, Ecran) then stChampsCompl := stChampsCompl + ',GA2_STATART' + stIndice ;
    end ;
  end ;

  // Paramétrage des libellés des tables libres article et dépôt
  if (GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'GA_LIBREART', 10, '') = 0) then SetControlVisible('PTABLESARTICLE', False) ;
  if (GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'YTC_TABLELIBRETIERS', 10, '') = 0) then SetControlVisible('PTABLESTIERS', False) ;
  // Mise en forme des libellés des dates, booléans libres et montants libres article et tiers
  Nbr := 0;
  if (GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'GA_VALLIBRE', 3, '_') = 0) then SetControlVisible('GB_VAL', False) else inc(Nbr) ;
  if (GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'GA_DATELIBRE', 3, '_') = 0) then SetControlVisible('GB_DATE', False) else inc(Nbr) ;
  if (GCMAJChampLibre (TForm (Ecran), False, 'BOOL', 'GA_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GB_BOOL', False) else inc(Nbr) ;
  {$IFNDEF CCS3}
  if (Abs(Nbr) = 0) then
  {$ENDIF}
  SetControlVisible ('PZONESARTICLE',False) ;

  Nbr := 0;
  if (GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_VALLIBRE', 3, '_') = 0) then SetControlVisible('GB_VAL1', False) else inc(Nbr) ;
  if (GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_DATELIBRE', 3, '_') = 0) then SetControlVisible('GB_DATE1', False) else inc(Nbr) ;
  if (GCMAJChampLibre (TForm (Ecran), False, 'BOOL', 'YTC_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GB_BOOL1', False) else inc(Nbr) ;
  {$IFNDEF CCS3}
  if (Abs(Nbr) = 0) then
  {$ENDIF}
  SetControlVisible('PZONESTIERS',False) ;
if TypeAction='RESERVE' then
   FTV.Caption := 'Modification des réservations'
   else FTV.Caption := 'Modification des affectations' ;
UpdateCaption (FTV) ;
end ;

procedure TOF_STATAFFCDECLIENT.OnUpdate ;
Var Q: TQuery ;
CodeArticle,SQL: String ;
i : Integer ;
begin
  Inherited ;
CodeArticle := '' ;
Q := OpenSQL(FTV.stSQL,True) ;
FTV.LaTOB:=TOB.Create('_AFFECTATION',nil,-1) ;
FTV.LaTOB.LoadDetailDB('les lignes','','',Q,False,True,-1,0) ;
Ferme(Q) ;
For i:=0 to FTV.LaTOB.Detail.Count-1 do
  begin
  if CodeArticle='' then
  CodeArticle :='"'+ FTV.LaTOB.Detail[i].GetValue('GEL_CODEARTICLE')+'"'
  else CodeArticle := CodeArticle + ',"' + FTV.LaTOB.Detail[i].GetValue('GEL_CODEARTICLE')+'"' ;
  end ;
If CodeArticle<>'' then
  begin
  SQL := 'SELECT * FROM AFFCDELIGNE WHERE GEL_CODEARTICLE IN ('+CodeArticle+')' ;
  Q := OpenSQL(SQL, True) ;
  TOBAffect:=TOB.Create('_Les affectations',nil,-1) ;
  TOBAffect.LoadDetailDB('AFFCDELIGNE','','',Q,False,True,-1,0) ;
  Ferme(Q) ;
  end ;
end ;

procedure TOF_STATAFFCDECLIENT.OnClose ;
Begin
  Inherited ;
FTV.LaTOB.Free ;
end ;

procedure TOF_STATAFFCDECLIENT.TVOnDblClick(Sender: TObject );
begin 
If FTV.TV.AsString[FTV.TV.ColIndex('GEL_TYPEDIM'), FTV.TV.CurrentRow] ='GEN' then TraiterArticleDimensionne
   else TraiterArticleUnique ;
end ;

procedure TOF_STATAFFCDECLIENT.TraiterArticleDimensionne ;
Var ValAction: String ;
begin
ValAction:='SAISIE' ;
RemplirTOBDim(TV1.AsString[FTV.TV.ColIndex('GEL_CODEARTICLE'), FTV.TV.CurrentRow],TV1.AsInteger[FTV.TV.ColIndex('GEL_NUMERO'), FTV.TV.CurrentRow]) ;
TheTOB := TOBDim ;
AglLanceFiche ('MBO','SAISIEDIMAFFCDE','','', 'GA_CODEARTICLE='+FTV.TV.AsString[FTV.TV.ColIndex('GEL_CODEARTICLE'), FTV.TV.CurrentRow]+';ACTION='+ValAction+';CHAMP= ;TYPEPARAM=ACC;DEPOT='+FTV.TV.AsString[FTV.TV.ColIndex('GEA_DEPOT'), FTV.TV.CurrentRow]+'') ;
if (TheTOB<>Nil) And (TheTOB.Detail.Count>0) then TOBDim:=TheTOB else TOBDim.ClearDetail ;
TOBDimToTV(TV1.AsInteger[FTV.TV.ColIndex('GEL_NUMERO'), FTV.TV.CurrentRow]) ;
end ;

procedure TOF_STATAFFCDECLIENT.TOBDimToTV(Commande: Integer) ;
Var i: integer ;
TOBA : TOB ;
TotQteReservee,TotQteAffectee : Double ;
CodeArticle : String ;
begin
for i:=0 to TOBDim.Detail.count-1 do
  begin
  TOBA:=TobAffect.FindFirst(['GEL_ARTICLE','GEL_NUMERO'],[TOBDim.Detail[i].GetValue('GA_ARTICLE'),Commande],False) ;
  if TOBA<>nil then
    begin
    TOBA.PutValue('GEL_QTERESERVEE',TOBDim.Detail[i].GetValue('GEL_QTERESERVEE')) ;
    TOBA.PutValue('GEL_QTEAFFECTEE',TOBDim.Detail[i].GetValue('GEL_QTEAFFECTEE')) ;
    TotQteReservee := TotQteReservee + TOBDim.Detail[i].GetValue('GEL_QTERESERVEE') ;
    TotQteAffectee := TotQteAffectee + TOBDim.Detail[i].GetValue('GEL_QTEAFFECTEE') ;
    CodeArticle := Trim(copy(TOBDim.Detail[i].GetValue('GA_ARTICLE'),1,18)) ;
    end ;
  end ;
// Mise à jour ligne générique
TOBA := TobAffect.FindFirst(['GEL_CODEARTICLE','GEL_NUMERO','GEL_TYPEDIM'],[CodeArticle,Commande,'GEN'],False) ;
if TOBA <> nil then
  begin
  TOBA.PutValue('GEL_QTERESERVEE',TotQteReservee) ;
  TOBA.PutValue('GEL_QTEAFFECTEE',TotQteAffectee) ;
  end ;
TV1.SetValueCell(TV1.ColIndex('GEL_QTERESERVEE'),TV1.CurrentROW,TotQteReservee);
TV1.SetValueCell(TV1.ColIndex('GEL_QTEAFFECTEE'),TV1.CurrentROW,TotQteAffectee);
//MiseAJourTable ;
TOBAffect.UpdateDB ;
end ;

procedure TOF_STATAFFCDECLIENT.TraiterArticleUnique ;
Var QteALivrer, QteReservee, QteAffectee: Double ;
Tiers,Article,QteModifie :String ;
Begin
Tiers:= TV1.AsString[TV1.ColIndex('GEL_TIERS'), TV1.CurrentRow] ;
Article:= TV1.AsString[TV1.ColIndex('GEL_CODEARTICLE'), TV1.CurrentRow] ;
QteALivrer:= TV1.AsDouble[TV1.ColIndex('GEL_QTEALIVRER'), TV1.CurrentRow] ;
QteReservee:= TV1.AsDouble[TV1.ColIndex('GEL_QTERESERVEE'), TV1.CurrentRow] ;
QteAffectee:= TV1.AsDouble[TV1.ColIndex('GEL_QTEAFFECTEE'), TV1.CurrentRow] ;
QteModifie:=AGLLanceFiche('MBO','SAISIEAFFCDE','','','QteALivrer='+FloatToStr(QteALivrer)+';QteReservee='+FloatToStr(QteReservee)+';QteAffectee='+FloatToStr(QteAffectee)+';Article='+Article+';Tiers='+Tiers+';TypeAction='+TypeAction) ;
QteReservee:=StrToFloat(ReadTokenSt(QteModifie)) ;
QteAffectee:=StrToFloat(QteModifie) ;
TV1.SetValueCell(TV1.ColIndex('GEL_QTERESERVEE'),TV1.CurrentROW,QteReservee);
TV1.SetValueCell(TV1.ColIndex('GEL_QTEAFFECTEE'),TV1.CurrentROW,QteAffectee);
MiseAJourTable ;
end ;  

// modif 02/08/2001
Procedure TOF_STATAFFCDECLIENT.RemplirTOBDim ( CodeArticle : String; Commande :Integer ) ;
Var TOBD,TOBA : TOB ;
    Q: TQuery ;
    QteReservee , QteAffectee , QteALivrer: Double ;
BEGIN
TOBDim.ClearDetail ;
Q:=OpenSQL('Select GA_Article from Article where GA_CodeArticle="'+CodeArticle+'" AND GA_STATUTART="DIM" order by GA_ARTICLE' ,True) ;
While not Q.EOF do
  begin
  TOBA := TOBAffect.findfirst(['GEL_ARTICLE','GEL_NUMERO'],[Q.FindField('GA_ARTICLE').AsString,Commande],False) ;
  if TOBA<>nil then
    begin
    TOBD:=TOB.Create('',TOBDim,-1) ;
    TOBD.AddChampSup('GA_ARTICLE',False)  ; TOBD.PutValue('GA_ARTICLE',Q.FindField('GA_ARTICLE').AsString) ;
    TOBD.AddChampSup('GEL_QTEALIVRER',False)  ; TOBD.PutValue('GEL_QTEALIVRER',TOBA.GetValue('GEL_QTEALIVRER')) ;
    TOBD.AddChampSup('GEL_QTERESERVEE',False)  ; TOBD.PutValue('GEL_QTERESERVEE',TOBA.GetValue('GEL_QTERESERVEE')) ;
    TOBD.AddChampSup('GEL_QTEAFFECTEE',False)  ; TOBD.PutValue('GEL_QTEAFFECTEE',TOBA.GetValue('GEL_QTEAFFECTEE')) ;
    end ;
  Q.Next ;
  end ;
ferme(Q) ;
//TheTOB:=TOBDim ;
END ;

procedure TOF_STATAFFCDECLIENT.MiseAJourTable ;
Var QteALivrer, QteReservee, QteAffectee: Double ;
CodeAff,Nature,Souche,QteModifie,numligne: String ;
numero,indiceg: integer ;
SQL: String ;
begin
QteReservee:= TV1.AsDouble[TV1.ColIndex('GEL_QTERESERVEE'), TV1.CurrentRow] ;
QteAffectee:= TV1.AsDouble[TV1.ColIndex('GEL_QTEAFFECTEE'), TV1.CurrentRow] ;
CodeAff:= TV1.AsString[TV1.ColIndex('GEA_CODEAFF'), TV1.CurrentRow] ;
Nature:= TV1.AsString[TV1.ColIndex('GEL_NATUREPIECEG'), TV1.CurrentRow] ;
Souche:= TV1.AsString[TV1.ColIndex('GEL_SOUCHE'), TV1.CurrentRow] ;
numero:= TV1.AsInteger[TV1.ColIndex('GEL_NUMERO'), TV1.CurrentRow] ;
indiceg:= TV1.AsInteger[TV1.ColIndex('GEL_INDICEG'), TV1.CurrentRow] ;
numligne:= TV1.AsString[TV1.ColIndex('GEL_NUMLIGNE'), TV1.CurrentRow] ;
SQL:='UPDATE AFFCDELIGNE SET GEL_QTERESERVEE="'+FloatToStr(QteReservee)+'", GEL_QTEAFFECTEE="'+FloatToStr(QteAffectee)+'", '+
     'GEL_CREERPAR="MAN", GEL_MODIFIABLE="-"  WHERE '+
     'GEL_CODEAFF="'+CodeAff+'" AND GEL_NATUREPIECEG="'+Nature+'" AND GEL_SOUCHE="'+Souche+'" AND '+
     'GEL_NUMERO="'+IntToStr(numero)+'" AND GEL_INDICEG="'+IntToStr(indiceg)+'" AND GEL_NUMLIGNE="'+numligne+'" AND (GEL_TYPEDIM="GEN" OR GEL_TYPEDIM="NOR")' ;
ExecuteSQL(SQL) ;
end ;




Initialization
  registerclasses ( [ TOF_STATAFFCDECLIENT ] ) ;
end.
