{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 19/02/2010
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTCHANGEVERS ()
Mots clefs ... : TOF;BTCHANGEVERS
*****************************************************************}
Unit BTCHANGEVERS_TOF ;

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
     uTob, 
     forms,
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF,
     vierge,
     UVersionParc,
     AglInit ;

Type
  TOF_BTCHANGEVERS = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	TOBEch : TOB;
    procedure CreerVersions;
    procedure ConstitueRequete;
  end ;

Implementation

procedure TOF_BTCHANGEVERS.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTCHANGEVERS.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTCHANGEVERS.OnUpdate ;
begin
  Inherited ;
  if Thedit(GetControl('CODEVERSION')).Text = '' then
  begin
  	PgiInfo ('veuillez renseigner la version');
    Thedit(GetControl('CODEVERSION')).SetFocus;
  	TFVierge(Ecran).ModalResult := 0;
		exit;
  end;
  if THCheckbox (GetControl('CREERVERSION')).checked then
  begin
    if ThValComboBox(GetControl('BVP_TYPEVERSION')).Value  = '' then
    begin
      PgiInfo ('veuillez renseigner le type de version');
    	ThValComboBox(GetControl('BVP_TYPEVERSION')).SetFocus;
      TFVierge(Ecran).ModalResult := 0;
      exit;
    end;
    if ThEdit(GetControl('BVP_LIBELLE')).text  = '' then
    begin
      PgiInfo ('veuillez renseigner le libellé de cette version');
      TFVierge(Ecran).ModalResult := 0;
      ThEdit(GetControl('BVP_LIBELLE')).SetFocus;
      exit;
    end;
    CreerVersions;
  end;
  TOBECH.putValue('RETOUR',Thedit(GetControl('CODEVERSION')).Text);

  TheTOB := TOBech;
end ;

procedure TOF_BTCHANGEVERS.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTCHANGEVERS.OnArgument (S : String ) ;
begin
  Inherited ;
  TOBEch := LaTOB;
  if TOBEch.getValue('FORCE')='X' then
  begin
		THCheckbox (GetControl('CREERVERSION')).checked := true;
		THCheckbox (GetControl('CREERVERSION')).enabled := false;
  end;
  ConstitueRequete;
end ;

procedure TOF_BTCHANGEVERS.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTCHANGEVERS.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTCHANGEVERS.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTCHANGEVERS.ConstitueRequete;
begin
end;

procedure TOF_BTCHANGEVERS.CreerVersions;
var VersionP : TVersionParc;
		TOBV : TOB;
    Articles,Article : string;
    QQ : TQuery;
    version : string;
begin
	version := Thedit(GetControl('CODEVERSION')).Text;
	versionP := TversionParc.create;
  versionP.ModeGestion := taCreat;
  TOBV := versionP.Nouveau;
  TRY
    Articles := TOBEch.GetValue('ARTICLES');
    repeat
    	Article := READTOKENST(Articles);
      if Article <> '' then
      begin
      	if not versionP.Existe(Article,version) then
        begin
          QQ := OpenSql ('SELECT GA_ARTICLE FROM ARTICLE WHERE GA_CODEARTICLE="'+Article+'"',true,1,'',true);
          if not QQ.eof then
          begin
            TOBV.putValue('BVP_ARTICLE',QQ.findField('GA_ARTICLE').asstring);
            TOBV.putValue('BVP_CODEARTICLE',Article);
            TOBV.putValue('BVP_CODEVERSION',version);
            TOBV.putValue('BVP_LIBELLE',ThEdit(GetControl('BVP_LIBELLE')).text);
            TOBV.putValue('BVP_TYPEVERSION',ThValComboBox(GetControl('BVP_TYPEVERSION')).Value);
            TOBV.putValue('BVP_ACTIVE','X');
            versionP.Ecrire(TOBV);
          end;
          ferme (QQ);
        end;
      end;
    until Articles = '';
  FINALLY
    TOBV.free;
    FreeAndNil(versionP);
  END;
end;

Initialization
  registerclasses ( [ TOF_BTCHANGEVERS ] ) ; 
end.
