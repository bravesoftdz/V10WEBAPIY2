unit UTofGCCATALOGU_MUL;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOF, HDimension,UTOM,AGLInit,
      Utob,HDB,Messages,HStatus,
{$IFDEF EAGLCLIENT}
      emul, MaineAGL,
{$ELSE}
      UTofGcModifLot,mul, db,DBGrids, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
{$ENDIF}
{$IFDEF BTP}
			paramsoc,
{$ENDIF}
      M3VM , M3FP,UtilArticle,Formule;

Type

    TOF_GCCATALOGU_MUL = Class (TOF)
        GCA_ARTICLE : THedit;
        GCA_ARTICLE_ : THedit;
        procedure OnLoad ; override ;
        procedure OnArgument (Arguments : String) ; override;
        procedure OnDisplay ; override;
        procedure OnUpdate; override;
        procedure GCA_ARTICLEOnExit(Sender : TObject) ;
        procedure GCA_ARTICLE_OnExit(Sender : TObject) ;
    end;

implementation

procedure TOF_GCCATALOGU_MUL.OnArgument (Arguments : String) ;
var stTiers, stArg, Critere, ChampMul : string;
    PosEgal : integer;
begin

if TControl(Ecran.FindComponent('TYPEACTION'))<>Nil then
   SetControlText('TYPEACTION','ACTION=MODIFICATION'); //Initialisation à modification
stArg:=Arguments; //Test si consultation
Repeat
    Critere := UpperCase(Trim(ReadTokenSt(stArg)));
    if Critere<>'' then
        begin
        PosEgal := pos('=',Critere);
        if PosEgal<>0 then ChampMul := copy(Critere,1,PosEgal-1);
           if ChampMul='ACTION' then SetControlText('TYPEACTION',Critere);
        end;
    until Critere='';
if TControl(Ecran.FindComponent('TYPEACTION'))<>Nil then
   SetControlVisible('bInsert',StringToAction(GetControlText('TYPEACTION'))<>taConsult);

if Ecran.Name = 'GCCATALOGUE_FOU' then
    begin
    stTiers := ReadTokenSt (Arguments);
    SetControlText ('XX_WHERE_FOUR', 'GCA_TIERS="' + stTiers + '"');
    Ecran.Caption := Ecran.Caption + ': ' + stTiers + '  ' + RechDom ('GCTIERSFOURN', stTiers, False);
    SetControlText ('ACTION', ReadTokenSt (Arguments));
    end;

GCA_ARTICLE := THEdit(GEtCOntrol('GCA_ARTICLE'));
GCA_ARTICLE.OnExit:=GCA_ARTICLEOnExit;
GCA_ARTICLE_ := THEdit(GEtCOntrol('GCA_ARTICLE_'));
GCA_ARTICLE_.OnExit:=GCA_ARTICLE_OnExit;

end;

procedure TOF_GCCATALOGU_MUL.OnDisplay;
begin
  inherited;
end;

procedure TOF_GCCATALOGU_MUL.GCA_ARTICLEOnExit(Sender : TObject) ;
begin
   SetControlText('GCA_ARTICLE',CodeArticleUnique(GetControlText('GCA_ARTICLE'),'','','','',''));
end;

procedure TOF_GCCATALOGU_MUL.GCA_ARTICLE_OnExit(Sender : TObject) ;
begin
   SetControlText('GCA_ARTICLE_',CodeArticleUnique(GetControlText('GCA_ARTICLE_'),'','','','',''));
end;

Procedure TOF_GCCATALOGU_MUL.OnLoad  ;
var  F : TFMul ;
    CC:TCheckBox ;
BEGIN
inherited ;
if not (Ecran is TFMul) then exit;
F:=TFMul(Ecran) ;
CC:= TCheckBox(TFMul(F).FindComponent('_ARTICLENONREF')) ;
If CC.State=cbChecked then
   Begin
   SetControlText('GCA_ARTICLE','') ;
   SetControlEnabled('GCA_ARTICLE',False) ;
   SetControlText('XX_WHERE','GCA_ARTICLE="" or GCA_ARTICLE is null') ;
   end ;
If CC.State=cbUnChecked then
   Begin
   SetControlEnabled('GCA_ARTICLE',True) ;
   SetControlText('XX_WHERE','GCA_ARTICLE<>""');
   End ;
If CC.State=cbGrayed then
   Begin
   SetControlEnabled('GCA_ARTICLE',True) ;
   SetControlText('XX_WHERE','');
   End ;
END;

procedure AGLArticleNonReference (parms: array of variant; nb: integer );
var  F : TFMul ;
begin
F := TFMul (Longint (Parms[0]));
if TCheckBox (F.FindComponent ('_ARTICLENONREF')).State = cbChecked then
    begin
    with THEdit(F.FindComponent('GCA_ARTICLE')) do
        begin
        Text := '';
        Enabled :=  False;
        end;
    THEdit(F.FindComponent('XX_WHERE')).Text := 'GCA_ARTICLE="" or GCA_ARTICLE is null';
    end;
if TCheckBox (F.FindComponent ('_ARTICLENONREF')).State = cbUnChecked then
    begin
    THEdit(F.FindComponent('GCA_ARTICLE')).Enabled :=  True;
    THEdit(F.FindComponent('XX_WHERE')).Text := 'GCA_ARTICLE<>""';
    end;
if TCheckBox(F.FindComponent('_ARTICLENONREF')).State = cbGrayed then
    begin
    THEdit(F.FindComponent('GCA_ARTICLE')).Enabled := True;
    THEdit(F.FindComponent('XX_WHERE')).Text := '';
    end;
end;

procedure TOF_GCCATALOGU_MUL.OnUpdate;
begin
  inherited;
end;

Initialization
registerclasses([TOF_GCCATALOGU_MUL]);
RegisterAglProc( 'ArticleNonReference', TRUE , 1, AGLArticleNonReference);
end.
