{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 12/05/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : DECISIONACHNART ()
Mots clefs ... : TOF;DECISIONACHNART
*****************************************************************}
Unit DECISIONACHNART_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
		 fe_main,
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
		 maineagl,
     eMul, 
     uTob, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1,
     HMsgBox,
     AglInit,
     UTOB,
     Vierge,
     UtilArticle,
     FactUtil,
     UTOF ;

Type
  TOF_DECISIONACHNART = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	TOBresult : TOB;
    Fournprinc : string;
    CodeArticle,Article : THEdit;
    Depot : THValComboBox;
    TLibArticle : THLabel;
    procedure CodeArticleElipsisClick (Sender : TObject);
    procedure DepotChange (Sender : TObject);
    function VerifCodeArticle : boolean;
  end ;


function GetArticle (var Article,Depot,fournprinc : string) : boolean;

Implementation


function GetArticle (var Article,Depot,fournprinc : string) : boolean;
var TOBparam : TOB;
begin
	result := false;
  //
  TOBParam := TOB.Create ('UNE TOB', nil,-1);
  TOBParam.AddChampSupValeur ('RESULT',-1);
  TOBParam.AddChampSupValeur ('ARTICLE','');
  TOBParam.AddChampSupValeur ('DEPOT','');
  TOBParam.AddChampSupValeur ('FOURNPRINC','');
  TheTOB := TOBParam;
  AglLanceFiche ('BTP','BTDECISIONACHNART','','','ACTION=CREATION');
  if TheTOB <> nil then
  begin
    if TheTOB.GetValue('RESULT') = 0 then
    begin
    	Article := TheTOB.getValue('ARTICLE');
    	Depot := TheTOB.getValue('DEPOT');
    	Fournprinc := TheTOB.getValue('FOURNPRINC');
      result := true;
    end;
  end;
  TheTOB := nil;
  TOBParam.free;
end;


procedure TOF_DECISIONACHNART.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_DECISIONACHNART.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_DECISIONACHNART.OnUpdate ;
begin
  Inherited ;
  if not VerifCodeArticle then
  begin
  	SetFocusControl ('CODEARTICLE');
    TFVierge(Ecran).ModalResult:=0;
    Exit;
  end;
  if Depot.value = '' then
  begin
  	SetFocusControl ('DEPOT');
    TFVierge(Ecran).ModalResult:=0;
    Exit;
  end;
  TOBresult.putValue('ARTICLE',ARTICLE.text);
  TOBresult.putValue('DEPOT',DEPOT.Value);
  TOBresult.putValue('FOURNPRINC',Fournprinc);
  TOBresult.putValue('RESULT',0);
  TheTOB := TOBResult;
end ;

procedure TOF_DECISIONACHNART.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_DECISIONACHNART.OnArgument (S : String ) ;
begin
  Inherited ;
  TOBresult := LaTOB;
  CodeArticle := THEdit(GetControl('CODEARTICLE'));
  Article := THEdit(GetControl('ARTICLE'));
  Depot := THValComboBox (GEtcontrol('DEPOT'));
  CodeArticle.OnElipsisClick := CodeArticleElipsisClick;
  Depot.OnChange := DepotChange;
end ;

procedure TOF_DECISIONACHNART.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_DECISIONACHNART.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_DECISIONACHNART.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_DECISIONACHNART.CodeArticleElipsisClick(Sender: TObject);
var SWhere,StChamps : string;
begin
	sWhere := 'AND ((GA_TYPEARTICLE="MAR") OR (GA_TYPEARTICLE="ARP") OR (GA_TYPEARTICLE="PRE"))';
  if CodeArticle.Text <> '' then StChamps := 'GA_CODEARTICLE=' + Trim(Copy(CodeArticle.Text, 1, 18)) + ';XX_WHERE=' + sWhere
														else StChamps := 'XX_WHERE=' + sWhere;

	Article.Text  := AGLLanceFiche('BTP', 'BTARTICLE_RECH', '', '', StChamps+';RECHERCHEARTICLE');
  CodeArticle.text := trim(Copy(Article.text,1,18));
  VerifCodeArticle;
end;

procedure TOF_DECISIONACHNART.DepotChange(Sender: TObject);
begin

end;

function TOF_DECISIONACHNART.VerifCodeArticle : boolean;
var QQ : TQUery;
begin
	result := false;
  if ((CodeArticle.text <> '') and (Article.text = '')) or (trim(Copy(Article.text,1,18))<>CodeArticle.text) then
  begin
  	Article.text := CodeArticleUnique2 (CodeArticle.text,'');
  end;
  QQ := OpenSql ('SELECT GA_STATUTART,GA_FOURNPRINC FROM ARTICLE WHERE GA_ARTICLE="'+Article.text+'"',true);
  if not QQ.eof then
  begin
    if QQ.FindField('GA_STATUTART').AsString  ='GEN' then
    begin
    	Article.text := SelectUneDimension (Article.text);
    end;
    if Article.text <> '' then
    begin
  	CodeArticle.text := trim(Copy(Article.text,1,18));
    FournPrinc := QQ.findField('GA_FOURNPRINC').AsString;
    result := true;
    end;
  end else
  begin
  	PgiBox(TraduireMemoire('Article inexistant..'));
  end;
  ferme (QQ);
end;

Initialization
  registerclasses ( [ TOF_DECISIONACHNART ] ) ;
end.

