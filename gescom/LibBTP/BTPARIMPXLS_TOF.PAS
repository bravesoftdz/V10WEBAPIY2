{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 08/06/2016
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTPARIMPXLS ()
Mots clefs ... : TOF;BTPARIMPXLS
*****************************************************************}
Unit BTPARIMPXLS_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
     FE_Main,
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
     HTB97,
     Hpanel,
     Aglinit,
     UTOF ;

Type
  TOF_BTPARIMPXLS = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    TOBPAR : TOB;
    procedure SetEcran;
    procedure SetLibelles;
    procedure SetEvents;
    procedure RB2Click (sender : Tobject);
    procedure RB1Click (sender : Tobject);
    procedure RechOuvrage (sender : Tobject);
    procedure RechMarchandise (sender : Tobject);
    procedure RechMOint(sender : Tobject);
    procedure RechInterim(sender : Tobject);
    procedure RechLocation(sender : Tobject);
    procedure RechMateriel(sender : Tobject);
    procedure RechOutillage(sender : Tobject);
    procedure RechSousTrait(sender : Tobject);
    //
    procedure CodeOuvExit(sender : Tobject);
    procedure CodeMarExit(sender : Tobject);
    procedure CodeMoExit(sender : Tobject);
    procedure CodeIntExit(sender : Tobject);
    procedure CodeLocExit(sender : Tobject);
    procedure CodeMatExit(sender : Tobject);
    procedure CodeOutExit(sender : Tobject);
    procedure CodeSTExit(sender : Tobject);
    //
    procedure ReinitArticles;
    function GetLibelleArticle(Article : string) : string;
    function RechPrestation (Typeressource,Article : string) : string;
    function RechArticle (TypeArt,Article : string): string;
    function ControleSaisie : boolean;
    function ExistePrestation(TypeRessource: string;Zone: Thedit): boolean;
    function DecomposeSql(TypeArt: string): string;
    function ExisteArticle(TypeArticle: string; Zone: Thedit): boolean;
  end ;

Implementation

procedure TOF_BTPARIMPXLS.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTPARIMPXLS.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTPARIMPXLS.OnUpdate ;
begin
  Inherited ;
  if not ControleSaisie then
  begin
    ecran.ModalResult := mrNone;
    exit;
  end;
  TOBPar.GetEcran(ecran); 
  TOBPAR.SetAllModifie(true); 
  TOBPAR.InsertOrUpdateDB(false); 
end ;

procedure TOF_BTPARIMPXLS.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTPARIMPXLS.OnArgument (S : String ) ;
var QQ : TQuery;
begin
  Inherited ;
  TOBPAR := TOB.create('BPARAMIMPXLS',nil,-1);
  QQ := OpenSql ('SELECT * FROM BPARAMIMPXLS',true,1,'',true);
  TOBPAR.SelectDB('',QQ);
  ferme (QQ);
  //
  ReinitArticles;
  SetEcran;
  SetEvents;
end ;

procedure TOF_BTPARIMPXLS.OnClose ;
begin
  TOBPAR.free;
  Inherited ;
end ;

procedure TOF_BTPARIMPXLS.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTPARIMPXLS.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTPARIMPXLS.SetEcran;
begin
  TOBPar.PutEcran(ecran);
  if TOBPAR.GetInteger('B01_NUMPARAM')=2 then
  begin
    TRadioButton (GetControl('RB2')).Checked := true;
    THpanel (GetControl('P2')).enabled := true;
    SetLibelles;
  end else
  begin
    TRadioButton (GetControl('RB1')).Checked := true;
    THpanel (GetControl('P2')).enabled := false;
  end;
end;

procedure TOF_BTPARIMPXLS.SetLibelles;
begin
  if GetControlText('B01_CODEOUV') <> '' then
  begin
    SetControlCaption('D0',GetLibelleArticle(GetControlText('B01_CODEOUV')));
  end;
  if GetControlText('B01_CODEMAR') <> '' then
  begin
    SetControlCaption('D1',GetLibelleArticle(GetControlText('B01_CODEMAR')));
  end;
  if GetControlText('B01_CODEMO') <> '' then
  begin
    SetControlCaption('D2',GetLibelleArticle(GetControlText('B01_CODEMO')));
  end;
  if GetControlText('B01_CODEINTERIM') <> '' then
  begin
    SetControlCaption('D3',GetLibelleArticle(GetControlText('B01_CODEINTERIM')));
  end;
  if GetControlText('B01_CODELOCATION') <> '' then
  begin
    SetControlCaption('D4',GetLibelleArticle(GetControlText('B01_CODELOCATION')));
  end;
  if GetControlText('B01_CODEMAT') <> '' then
  begin
    SetControlCaption('D5',GetLibelleArticle(GetControlText('B01_CODEMAT')));
  end;
  if GetControlText('B01_CODEOUTIL') <> '' then
  begin
    SetControlCaption('D6',GetLibelleArticle(GetControlText('B01_CODEOUTIL')));
  end;
  if GetControlText('B01_CODESOUSTRAIT') <> '' then
  begin
    SetControlCaption('D7',GetLibelleArticle(GetControlText('B01_CODESOUSTRAIT')));
  end;
end;

procedure TOF_BTPARIMPXLS.SetEvents;
begin
  TRadioButton (GetControl('RB2')).OnClick := RB2Click;
  TRadioButton (GetControl('RB1')).OnClick := RB1Click;
  //
  THEdit (GetControl('B01_CODEOUV')).OnElipsisClick := RechOuvrage;
  THEdit (GetControl('B01_CODEMAR')).OnElipsisClick := RechMarchandise;
  THEdit (GetControl('B01_CODEMO')).OnElipsisClick := RechMOint;
  THEdit (GetControl('B01_CODEINTERIM')).OnElipsisClick := RechInterim;
  THEdit (GetControl('B01_CODELOCATION')).OnElipsisClick := RechLocation;
  THEdit (GetControl('B01_CODEMAT')).OnElipsisClick := RechMateriel;
  THEdit (GetControl('B01_CODEOUTIL')).OnElipsisClick := RechOutillage;
  THEdit (GetControl('B01_CODESOUSTRAIT')).OnElipsisClick := RechSousTrait;
  //
  THEdit (GetControl('B01_CODEOUV')).OnExit := CodeOuvExit;
  THEdit (GetControl('B01_CODEMAR')).OnExit := CodeMarExit;
  THEdit (GetControl('B01_CODEMO')).OnExit := CodeMoExit;
  THEdit (GetControl('B01_CODEINTERIM')).OnExit := CodeIntExit;
  THEdit (GetControl('B01_CODELOCATION')).OnExit := CodeLocExit;
  THEdit (GetControl('B01_CODEMAT')).OnExit := CodeMatExit;
  THEdit (GetControl('B01_CODEOUTIL')).OnExit := CodeOutExit;
  THEdit (GetControl('B01_CODESOUSTRAIT')).OnExit := CodeSTExit;
end;

procedure TOF_BTPARIMPXLS.RB1Click(sender: Tobject);
begin
  TOBPAr.setInteger('B01_NUMPARAM',1);
  THpanel (GetControl('P2')).enabled := true;
  ReinitArticles;
end;

procedure TOF_BTPARIMPXLS.RB2Click(sender: Tobject);
begin
  TOBPAr.setInteger('B01_NUMPARAM',2);
  THpanel (GetControl('P2')).enabled := true;
  SetLibelles;
end;

procedure TOF_BTPARIMPXLS.RechInterim(sender: Tobject);
var result : string;
begin
  result := RechPrestation ('INT',getControlText('B01_CODEINTERIM'));
  if result <> '' then
  begin
    SetControlText('B01_CODEINTERIM',result);
    SetControlCaption('D3',GetLibelleArticle(GetControlText('B01_CODEINTERIM')));
  end;
end;

procedure TOF_BTPARIMPXLS.RechLocation(sender: Tobject);
var result : string;
begin
  result := RechPrestation ('LOC',getControlText('B01_CODELOCATION'));
  if result <> '' then
  begin
    SetControlText('B01_CODELOCATION',result);
    SetControlCaption('D4',GetLibelleArticle(GetControlText('B01_CODELOCATION')));
  end;
end;

procedure TOF_BTPARIMPXLS.RechMarchandise(sender: Tobject);
var result : string;
begin
  result := RechArticle('MAR;ARP',getControlText('B01_CODEMAR'));
  if result <> '' then
  begin
    SetControlText('B01_CODEMAR',result);
    SetControlCaption('D1',GetLibelleArticle(GetControlText('B01_CODEMAR')));
  end;
end;

procedure TOF_BTPARIMPXLS.RechMateriel(sender: Tobject);
var result : string;
begin
  result := RechPrestation ('MAT',getControlText('B01_CODEMAT'));
  if result <> '' then
  begin
    SetControlText('B01_CODEMAT',result);
    SetControlCaption('D5',GetLibelleArticle(GetControlText('B01_CODEMAT')));
  end;
end;

procedure TOF_BTPARIMPXLS.RechMOint(sender: Tobject);
var result : string;
begin
  result := RechPrestation ('SAL',getControlText('B01_CODEMO'));
  if result <> '' then
  begin
    SetControlText('B01_CODEMO',result);
    SetControlCaption('D2',GetLibelleArticle(GetControlText('B01_CODEMO')));
  end;
end;

procedure TOF_BTPARIMPXLS.RechOutillage(sender: Tobject);
var result : string;
begin
  result := RechPrestation ('OUT',getControlText('B01_CODEOUTIL'));
  if result <> '' then
  begin
    SetControlText('B01_CODEOUTIL',result);
    SetControlCaption('D6',GetLibelleArticle(GetControlText('B01_CODEOUTIL')));
  end;
end;

procedure TOF_BTPARIMPXLS.RechOuvrage(sender: Tobject);
var result : string;
begin
  result := RechArticle('OUV',getControlText('B01_CODEOUV'));
  if result <> '' then
  begin
    SetControlText('B01_CODEOUV',result);
    SetControlCaption('D0',GetLibelleArticle(GetControlText('B01_CODEOUV')));
  end;
end;

procedure TOF_BTPARIMPXLS.RechSousTrait(sender: Tobject);
var result : string;
begin
  result := RechPrestation ('ST',getControlText('B01_CODESOUSTRAIT'));
  if result <> '' then
  begin
    SetControlText('B01_CODESOUSTRAIT',result);
    SetControlCaption('D7',GetLibelleArticle(GetControlText('B01_CODESOUSTRAIT')));
  end;
end;

function TOF_BTPARIMPXLS.GetLibelleArticle(Article: string): string;
var QQ : TQuery;
begin
  result := '';
  if Article = '' then exit;
  QQ := OpenSql ('SELECT GA_LIBELLE FROM ARTICLE WHERE GA_ARTICLE="'+Article+'"',true,1,'',true);
  if not QQ.eof then
  begin
    result := QQ.fields[0].AsString;
  end;
  ferme (QQ);
end;

procedure TOF_BTPARIMPXLS.ReinitArticles;
begin
  THlabel(GetControl('D0')).Caption := '';
  THlabel(GetControl('D1')).Caption := '';
  THlabel(GetControl('D2')).Caption := '';
  THlabel(GetControl('D3')).Caption := '';
  THlabel(GetControl('D4')).Caption := '';
  THlabel(GetControl('D5')).Caption := '';
  THlabel(GetControl('D6')).Caption := '';
  THlabel(GetControl('D7')).Caption := '';
  SetControlText('B01_CODEOUV','');
  SetControlText('B01_CODEMAR','');
  SetControlText('B01_CODEMO','');
  SetControlText('B01_CODEINTERIM','');
  SetControlText('B01_CODELOCATION','');
  SetControlText('B01_CODEMAT','');
  SetControlText('B01_CODEOUTIL','');
  SetControlText('B01_CODESOUSTRAIT','');
end;

function TOF_BTPARIMPXLS.DecomposeSql(TypeArt : string) : string;
var ZZ,ZZ1 : string;
    i : integer;
begin
  result := '';
  ZZ := TypeArt;
  i := 0;
  repeat
    ZZ1 := READTOKENST (ZZ);
    if ZZ1 <> '' then
    begin
      if result = '' then
      begin
        result := ' GA_TYPEARTICLE IN (';
      end;
      if I > 0 then result := result +',';
      result := result + '"'+ZZ1+'"';
      inc(i);
    end;
  until ZZ1 = '';
  if result <> '' then result := result + ')';
end;

function TOF_BTPARIMPXLS.RechArticle (TypeArt,Article : string): string;
Var  stchamps : String;
begin

  stChamps := 'XX_WHERE='+DecomposeSql(TypeArt)+';FIXEDTYPEART';
  result := AGLLanceFiche('BTP', 'BTARTICLE_RECH', 'GA_CODEARTICLE='+copy(Article,1,18), '', StChamps);
end;

function TOF_BTPARIMPXLS.RechPrestation(Typeressource,Article: string): string;
Var  stchamps : String;
begin

  stchamps := 'TYPERESSOURCE='+ Typeressource;
  stchamps := stchamps + ';GA_TYPEARTICLE=PRE';
  result := AGLLanceFiche('BTP', 'BTPREST_RECH', 'GA_CODEARTICLE='+copy(Article,1,18) , '', stchamps);
end;

procedure TOF_BTPARIMPXLS.CodeIntExit(sender: Tobject);
begin
  if not ExistePrestation ('INT',Thedit(GetControl('B01_CODEINTERIM'))) then
  begin
    PgiError ('Prestation Interimaire inexistante...');
    SetControlText('B01_CODEINTERIM',''); THEdit(getControl('B01_CODEINTERIM')).SetFocus;
    exit;
  end;
  SetControlCaption('D3',GetLibelleArticle(GetControlText('B01_CODEINTERIM')));
end;

procedure TOF_BTPARIMPXLS.CodeLocExit(sender: Tobject);
begin
  if not ExistePrestation ('LOC',THEdit(GetControl('B01_CODELOCATION'))) then
  begin
    PgiError ('Prestation Location inexistante...');
    SetControlText('B01_CODELOCATION',''); THEdit(getControl('B01_CODELOCATION')).SetFocus;
    exit;
  end;
  SetControlCaption('D4',GetLibelleArticle(GetControlText('B01_CODELOCATION')));
end;

procedure TOF_BTPARIMPXLS.CodeMarExit(sender: Tobject);
begin
  if not ExisteArticle ('MAR;ARP',ThEdit(GetControl('B01_CODEMAR'))) then
  begin
    PgiError ('Article inexistant...');
    SetControlText('B01_CODEMAR',''); THEdit(getControl('B01_CODEMAR')).SetFocus;
    exit;
  end;
  SetControlCaption('D1',GetLibelleArticle(GetControlText('B01_CODEMAR')));
end;

procedure TOF_BTPARIMPXLS.CodeMatExit(sender: Tobject);
begin
  if not ExistePrestation ('MAT',ThEdit(GetControl('B01_CODEMAT'))) then
  begin
    PgiError ('Prestation Matérielle inexistante...');
    SetControlText('B01_CODEMAT',''); THEdit(getControl('B01_CODEMAT')).SetFocus;
    exit;
  end;
  SetControlCaption('D5',GetLibelleArticle(GetControlText('B01_CODEMAT')));
end;

procedure TOF_BTPARIMPXLS.CodeMoExit(sender: Tobject);
begin
  if not ExistePrestation ('SAL',THEdit(GetControl('B01_CODEMO'))) then
  begin
    PgiError ('Prestation Main oeuvre interne inexistante...');
    SetControlText('B01_CODEMO',''); THEdit(getControl('B01_CODEMO')).SetFocus;
    exit;
  end;
  SetControlCaption('D2',GetLibelleArticle(GetControlText('B01_CODEMO')));
end;

procedure TOF_BTPARIMPXLS.CodeOutExit(sender: Tobject);
begin
  if not ExistePrestation ('OUT',THEdit(GetControl('B01_CODEOUTIL'))) then
  begin
    PgiError ('Prestation outillage inexistante...');
    SetControlText('B01_CODEOUTIL','');THEdit(getControl('B01_CODEOUTIL')).SetFocus;
    exit;
  end;
  SetControlCaption('D6',GetLibelleArticle(GetControlText('B01_CODEOUTIL')));
end;

procedure TOF_BTPARIMPXLS.CodeOuvExit(sender: Tobject);
begin
  if not ExisteArticle ('OUV',THEdit(GetControl('B01_CODEOUV'))) then
  begin
    PgiError ('Ouvrage inexistant...');
    SetControlText('B01_CODEOUV',''); THEdit(getControl('B01_CODEOUV')).SetFocus;
    exit;
  end;
  SetControlCaption('D0',GetLibelleArticle(GetControlText('B01_CODEOUV')));
end;

procedure TOF_BTPARIMPXLS.CodeSTExit(sender: Tobject);
begin
  if not ExistePrestation ('ST',THEdit(GetControl('B01_CODESOUSTRAIT'))) then
  begin
    PgiError ('Prestation de sous-traitance inexistante...');
    SetControlText('B01_CODESOUSTRAIT',''); THEdit(getControl('B01_CODESOUSTRAIT')).SetFocus;
    exit;
  end;
  SetControlCaption('D7',GetLibelleArticle(GetControlText('B01_CODESOUSTRAIT')));
end;

function TOF_BTPARIMPXLS.ControleSaisie: boolean;
begin
  result := true;
  if not ExisteArticle ('OUV',THEdit(GetControl('B01_CODEOUV'))) then
  begin
    PgiError ('Ouvrage inexistant...');
    SetControlText('B01_CODEOUV',''); THEdit(getControl('B01_CODEOUV')).SetFocus; 
    result := false;
    exit;
  end;
  if not ExisteArticle ('MAR;ARP',THEdit(GetControl('B01_CODEMAR'))) then
  begin
    PgiError ('Article inexistant...');
    SetControlText('B01_CODEMAR',''); THEdit(getControl('B01_CODEMAR')).SetFocus;
    result := false;
    exit;
  end;
  if not ExistePrestation ('SAL',THEdit(GetControl('B01_CODEMO'))) then
  begin
    PgiError ('Prestation Main oeuvre interne inexistante...');
    SetControlText('B01_CODEMO',''); THEdit(getControl('B01_CODEMO')).SetFocus;
    result := false;
    exit;
  end;
  if not ExistePrestation ('INT',THEdit(GetControl('B01_CODEINTERIM'))) then
  begin
    PgiError ('Prestation Interimaire inexistante...');
    SetControlText('B01_CODEINTERIM',''); THEdit(getControl('B01_CODEINTERIM')).SetFocus;
    result := false;
    exit;
  end;
  if not ExistePrestation ('LOC',THEdit(GetControl('B01_CODELOCATION'))) then
  begin
    PgiError ('Prestation Location inexistante...');
    SetControlText('B01_CODELOCATION',''); THEdit(getControl('B01_CODELOCATION')).SetFocus;
    result := false;
    exit;
  end;
  if not ExistePrestation ('MAT',THEdit(GetControl('B01_CODEMAT'))) then
  begin
    PgiError ('Prestation Matérielle inexistante...');
    SetControlText('B01_CODEMAT',''); THEdit(getControl('B01_CODEMAT')).SetFocus;
    result := false;
    exit;
  end;
  if not ExistePrestation ('OUT',THEdit(GetControl('B01_CODEOUTIL'))) then
  begin
    PgiError ('Prestation outillage inexistante...');
    SetControlText('B01_CODEOUTIL',''); THEdit(getControl('B01_CODEOUTIL')).SetFocus;
    result := false;
    exit;
  end;
  if not ExistePrestation ('ST',THEdit(GetControl('B01_CODESOUSTRAIT'))) then
  begin
    PgiError ('Prestation de sous-traitance inexistante...');
    SetControlText('B01_CODESOUSTRAIT','');THEdit(getControl('B01_CODESOUSTRAIT')).SetFocus;
    result := false;
    exit;
  end;
end;

function TOF_BTPARIMPXLS.ExisteArticle (TypeArticle: string; Zone : Thedit) : boolean;
var SQl : string;
    Ouvrage : string;
    QQ : TQuery;
begin
  result := true;
  Ouvrage := Zone.Text;
  if ouvrage = '' then exit;
  Sql := 'SELECT 1 FROM ARTICLE WHERE GA_ARTICLE="'+Ouvrage+'" AND '+DecomposeSql(TypeArticle);
  result := ExisteSql (Sql);
  if not result then
  begin
    Sql :='SELECT GA_ARTICLE FROM ARTICLE WHERE '+
          'GA_ARTICLE LIKE "'+Trim(Ouvrage)+'%" AND '+DecomposeSql(TypeArticle);
    QQ := OpenSql (Sql,true,-1,'',false);
    if not QQ.eof then
    begin
      if QQ.RecordCount = 1 then
      begin
        ThEdit(Zone).Text := QQ.fields[0].AsString;
        result := true;
      end;
    end;
    ferme (QQ);
  end;
end;

function TOF_BTPARIMPXLS.ExistePrestation(TypeRessource: string;Zone: Thedit): boolean;
var Sql : string;
    Article : string;
    QQ : Tquery;
begin
  result := true;
  Article := Zone.Text;
  if Article = '' then exit;
  Sql :='SELECT 1 FROM ARTICLE WHERE '+
       'GA_ARTICLE="'+article+'" AND GA_TYPEARTICLE="PRE" AND '+
       '(SELECT BNP_TYPERESSOURCE FROM NATUREPREST WHERE BNP_NATUREPRES=GA_NATUREPRES)="'+TypeRessource+'"';
  result := ExisteSql (Sql);
  if not result then
  begin
    Sql :='SELECT GA_ARTICLE FROM ARTICLE WHERE '+
         'GA_ARTICLE like "'+TRIM(article)+'%" AND GA_TYPEARTICLE="PRE" AND '+
         '(SELECT BNP_TYPERESSOURCE FROM NATUREPREST WHERE BNP_NATUREPRES=GA_NATUREPRES)="'+TypeRessource+'"';
    QQ := OpenSql (Sql,true,-1,'',false);
    if not QQ.eof then
    begin
      if QQ.RecordCount = 1 then
      begin
        ThEdit(Zone).Text := QQ.fields[0].AsString;
        result := true;
      end;
    end;
    ferme (QQ);
  end;
end;

Initialization
  registerclasses ( [ TOF_BTPARIMPXLS ] ) ;
end.
