{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 06/02/2010
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BPARCVERSION ()
Mots clefs ... : TOF;BPARCVERSION
*****************************************************************}
Unit BPARCVERSION_TOF ;

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
{$ENDIF}
		 vierge,
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     ExtCtrls,
     HEnt1,
     Hpanel,
     HMsgBox,
     UTOF,
     HTB97,
     HRichOLE,
     AglInit,
     Messages,
     windows,
     UversionParc,
     UtilArticle,
     AglInitGc
 ;

Type
  TOF_BPARCVERSION = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
	private
  	fID : integer;
  	fAction : TactionFiche;
    TCtrl : TVersionParc;
    TOBV : TOB;
    procedure BLOCNOTEExit (Sender : TObject);
  	procedure CodeArticleExit (Sender : TObject);
    procedure CodeVersionExit (Sender : Tobject);
    function getLibelleArticle (Article : string) : string;
    procedure LIBELLEExit (Sender : TObject);
    procedure RechArticle (Sender : TObject);
    procedure SetParams (Sparams : string);
    procedure SetEvents;
  end ;

Implementation

procedure TOF_BPARCVERSION.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BPARCVERSION.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BPARCVERSION.OnUpdate ;
begin
  Inherited ;
  TOBV.GetEcran(ecran);
  if TOBV.getValue('BVP_CODEARTICLE')='' then
  begin
  	PgiInfo('Veuillez renseigner l''élément parc');
    THEdit(GetControl('BVP_CODEARTICLE')).setfocus;
		TFVierge(Ecran).ModalResult := 0;
    exit;
  end
  else if TOBV.getValue('BVP_CODEVERSION')='' then
  begin
  	PgiInfo('Veuillez renseigner la version');
    THEdit(GetControl('BVP_CODEVERSION')).setfocus;
		TFVierge(Ecran).ModalResult := 0;
    exit;
  end
  else if TOBV.getValue('BVP_TYPEVERSION')='' then
  begin
  	PgiInfo('Veuillez renseigner le type de version');
    THValComboBox(GetControl('BVP_TYPEVERSION')).setfocus;
		TFVierge(Ecran).ModalResult := 0;
    exit;
  end;
  TCtrl.Ecrire(TOBV);
end ;

procedure TOF_BPARCVERSION.OnLoad ;
begin
  Inherited ;
  ThEdit(getControl('LIBELLEARTICLE')).Text :='';
  if faction <> taCreat then
  begin
  	if not TCtrl.lire (fID,TOBV) then
    begin
    	PgiError('La fiche n''existe pas');
      PostMessage(ecran.Handle, WM_CLOSE, 0, 0);
      exit;
    end;
    ThEdit(getControl('LIBELLEARTICLE')).Text := getLibelleArticle (TOBV.GetValue('BVP_ARTICLE'));
    ThEdit(getControl('BVP_CODEARTICLE')).ElipsisButton := false;
    ThPanel(getControl('PELTPARC')).enabled := false;
  end;
  TOBV.PutEcran(ecran);
  stringtorich( THricheditOle(getControl('BVP_BLOCNOTE')),TOBV.GetValue('BVP_BLOCNOTE'));
  SetEvents;
end ;

procedure TOF_BPARCVERSION.OnArgument (S : String ) ;
begin
  Inherited ;
  TCtrl := TVersionParc.create;
  TOBV := TCtrl.Nouveau;
  SetParams(S);
  if Faction = taCreat then
  begin
  	TToolBarButton97(GetControl('BDelete')).visible := false;
  end;
  TCtrl.ModeGestion := Faction;
end ;

procedure TOF_BPARCVERSION.OnClose ;
begin
	freeAndNil(TOBV);
  freeAndNil(TCTRL);
  Inherited ;
end ;

procedure TOF_BPARCVERSION.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BPARCVERSION.OnCancel () ;
begin
  Inherited ;
end ;


procedure TOF_BPARCVERSION.SetParams(Sparams: string);
var Critere,ChampMul,Valmul : string;
		X : integer;
begin
  repeat
    Critere := uppercase(Trim(ReadTokenSt(sParams)));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        ChampMul := copy(Critere, 1, x - 1);
        ValMul := copy(Critere, x + 1, length(Critere));
        if ChampMul = 'ACTION' then
        begin
        	fAction := StringToAction(critere);
        end else if champMul = 'ID' then
        begin
        	fId := StrtoInt(ValMul);
        end else if champMul = 'CODEARTICLE' then
        begin
        	TOBV.putValue('BVP_CODEARTICLE', ValMul);
        end;
      end;
    end;
  until sparams = '';
end;

procedure TOF_BPARCVERSION.SetEvents;
begin
	if faction = TaCreat then
  begin
  	Thedit(getControl('BVP_CODEARTICLE')).OnExit := CodeArticleExit;
    THEdit(GetControl('BVP_CODEARTICLE')).OnElipsisClick := RechArticle;
  end;
  Thedit(getControl('BVP_CODEVERSION')).OnExit := CodeVersionExit;
  Thedit(getControl('BVP_LIBELLE')).OnExit := LIBELLEExit;
  THricheditOle(getControl('BVP_BLOCNOTE')).OnExit := BLOCNOTEExit;
end;

procedure TOF_BPARCVERSION.BLOCNOTEExit(Sender: TObject);
begin
	if (Length(THricheditOle(getControl('BVP_BLOCNOTE')).Text ) <> 0) and (THricheditOle(getControl('BVP_BLOCNOTE')).text <> #$D#$A) then
  begin
    TOBV.PutValue('BVP_BLOCNOTE', RichToString(THricheditOle(getControl('BVP_BLOCNOTE'))));
  end else
  begin
  	TOBV.PutValue('BVP_BLOCNOTE','');
  end;
end;

procedure TOF_BPARCVERSION.CodeArticleExit(Sender: TObject);
var QQ : TQuery;
begin
  QQ := OpenSql ('SELECT GA_CODEARTICLE,GA_ARTICLE,GA_LIBELLE FROM ARTICLE WHERE GA_CODEARTICLE="'+
                 Thedit(getControl('BVP_CODEARTICLE')).Text+'" AND GA_TYPEARTICLE LIKE "PA%"',true,1,'',true);
  if not QQ.eof then
  begin
  	Thedit(getControl('BVP_CODEARTICLE')).text := QQ.findfield('GA_CODEARTICLE').asString;
  	Thedit(getControl('LIBELLEARTICLE')).text := QQ.findfield('GA_LIBELLE').asString;
    TOBV.putValue('BVP_CODEARTICLE',QQ.findfield('GA_CODEARTICLE').asString);
    TOBV.putValue('BVP_ARTICLE',QQ.findfield('GA_ARTICLE').asString);
  end else
  begin
  	PgiError ('Element parc inexistant');
    Thedit(getControl('BVP_CODEARTICLE')).SetFocus;
  end;
  ferme (QQ);
end;

procedure TOF_BPARCVERSION.CodeVersionExit(Sender: Tobject);
var QQ: TQuery;
begin
	if Thedit(getControl('BVP_CODEVERSION')).text = TOBV.getValue('BVP_CODEVERSION') then exit;
  QQ := OpenSql ('SELECT 	BVP_CODEVERSION FROM BVERSIONSPARC WHERE BVP_ARTICLE="'+
                 TOBV.GetValue('BVP_ARTICLE')+'" AND BVP_CODEVERSION ="'+Thedit(getControl('BVP_CODEVERSION')).text+'"',true,1,'',true);
  if QQ.eof then
  begin
  	TOBV.putValue('BVP_CODEVERSION',Thedit(getControl('BVP_CODEVERSION')).text);
  end else
  begin
  	PgiError ('Version Element parc existante');
    Thedit(getControl('BVP_CODEVERSION')).SetFocus;
  end;
  ferme (QQ);
end;

procedure TOF_BPARCVERSION.LIBELLEExit(Sender: TObject);
begin
	TOBV.putValue('BVP_LIBELLE',Thedit(getControl('BVP_LIBELLE')).text);
end;

procedure TOF_BPARCVERSION.RechArticle(Sender: TObject);
var stWhere,stRange, stFiche : string;
    ART : THCritMaskEdit;
begin
	ART := THCritMaskEdit.create(ecran);
  ART.Parent := ecran;
  ART.Top := Thedit(getControl('BVP_CODEARTICLE')).Top;
  ART.Left := Thedit(getControl('BVP_CODEARTICLE')).Left;
  ART.Width := 3; ART.Visible := False;
  ART.DataType := 'GCARTICLE';
  ART.Text := '';
  StRange := GetTypeArticlePARC;
  stWhere := 'BCP_ISVERSION=X';
  stFiche := 'BTARTPARC_RECH';
	DispatchRecherche (ART, 1, stWhere,stRange, stFiche);
  if ART.text <> '' then
  begin
  	Thedit(getControl('BVP_CODEARTICLE')).text := copy(ART.text,1,18);
  end;
  ART.free;
end;

function TOF_BPARCVERSION.getLibelleArticle(Article: string): string;
var QQ : TQuery;
begin
	result := '';
  QQ := OpenSql ('SELECT GA_LIBELLE FROM ARTICLE WHERE GA_ARTICLE="'+Article+'"',true,1,'',true);
  if not QQ.eof then result := QQ.findField('GA_LIBELLE').asString;
  ferme (QQ);
end;

Initialization
  registerclasses ( [ TOF_BPARCVERSION ] ) ; 
end.
