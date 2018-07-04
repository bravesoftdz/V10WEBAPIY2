{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 27/03/2001
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : PARBONI (PARBONI)
Mots clefs ... : TOM;PARBONI
*****************************************************************}
Unit UTOMPARBONI ;

Interface

Uses StdCtrls, Controls, Classes,  forms, sysutils, ComCtrls,
{$IFDEF EAGLCLIENT}
        Maineagl,  eFiche, eFichList,
{$ELSE}
   dbTables, db,FE_Main, Fiche, FichList,
{$ENDIF}
     HCtrls, HEnt1, HMsgBox, UTOM,   UTob,
     utilarticle,dicoaf,lookup
     ,AfUtilArticle;

Type
  TOM_PARBONI = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    end ;
    const
	// libellés des messages de la TOM
	TexteMsgParboni: array[1..8] of string 	= (
          {1}        'Le code article est invalide',
          {2}        'Incohérence dans les ordres de priorité',
          {3}        'Le sens doit être ''+''  ou  ''-''',
          {4}        'Le code article est obligatoire',
          {5}        'Vous devez renseigner au moins 2 choix',
          {6}        'L''ordre de priorité doit être renseigner',
          {7}        'L''ordre de priorité ne doit pas être renseigner',
          {8}        'Code ressource inconnu '
          );

Procedure AFLanceFiche_ParamBoni;
Implementation

procedure TOM_PARBONI.OnNewRecord ;
 Var QQ : TQuery ;
    IMax :integer ;
begin
  Inherited ;
  QQ := nil;
  try
    QQ:=OpenSQL('SELECT MAX(APB_NUMBONI) FROM PARBONI',TRUE);
    if Not QQ.EOF then Imax:=QQ.Fields[0].AsInteger+1 else iMax:=1;
  finally
    Ferme(QQ);              
  end;

  SetField('APB_NUMBONI',IMax);
  SetField('APB_NUMLIG',1);                       
end ;

procedure TOM_PARBONI.OnDeleteRecord ; 
begin
  Inherited ;
end ;

procedure TOM_PARBONI.OnUpdateRecord ;
Var za,zb,zc,wnb: Integer;
		zsens,dim,codeart,typart : string;
    zpro  : TCheckBox;
    zart : boolean;
    qq : TQuery ;

begin
  Inherited ;
//  controle zone  SENS

  zsens := GetField('APB_SENS');
  if (zsens<>'-') and (zsens<>'+') then
   Begin
  	LastError:=3; LastErrorMsg:=TraduitGa(TexteMsgParboni[LastError]);
    SetFocusControl('APB_SENS');
    Exit;
  End;

//  controle sur code Article
dim := '';
  // mcd 21/03/03 tout revu car accès vio en 4.2.0 et > sur zzart:=qqeof
if THEdit(GetControl('APB_CODEARTICLE')).Text<>'' then
  Begin
	CodeArt := CodeArticleUnique(THEdit(GetControl('APB_CODEARTICLE')).Text,dim,dim,dim,dim,dim);
  TypArt:='' ;
  QQ:=OpenSQL('SELECT GA_TYPEARTICLE FROM ARTICLE WHERE GA_ARTICLE="'+ CodeArt +'"',true);
  zart := qq.eof;
  if not QQ.EOF then
    Begin
    TypArt:=QQ.findField('GA_TYPEARTICLE').asString ;
    SetField('APB_TYPEARTICLE',TypArt);
    End;
  ferme(QQ);
  if (zart) then
    Begin
  	LastError:=1; LastErrorMsg:=TraduitGa(TexteMsgParboni[LastError]);
    SetFocusControl('APB_CODEARTICLE');
    Exit;
    End;
  End;
if THEdit(GetControl('APB_CODEARTICLE')).Text='' then
 Begin
  	LastError:=4; LastErrorMsg:=TraduitGa(TexteMsgParboni[LastError]);
    SetFocusControl('APB_CODEARTICLE');
    Exit;
  End;
  // Gestion du passage du CODEARTICLE en ARTICLE pour stocker
  if (THEdit(GetControl('APB_CODEARTICLE')).Text='') then
  	SetField('APB_ARTICLE', '')
  else
		SetField('APB_ARTICLE', CodeArticleUnique(CodeArt,'','','','','')) ;

  // CONTROLE RESSOURCE
   if (GetField('APB_RESSOURCE') <> '' ) then
     if Not(LookupValueExist(Getcontrol('APB_RESSOURCE'))) then
      Begin
      LastError:=8;
      LastErrorMsg:=TraduitGa(TexteMsgParboni[LastError]);
      SetFocusControl('APB_RESSOURCE');
      Exit;
      End;


    // Controle sur les zones de dispatch
  // on doit avoir au moins 2 choix
  wnb := 0;
  if (GetField('APB_RESSOURCE') <> '') then inc(wnb);
  if (GetField('APB_RESSAFF') <> '') then inc(wnb);
  zpro :=TCheckBox(GetControl('APB_PRORATA'));
  if zpro.checked then inc(wnb);

  if wnb < 2  then
  Begin
  	LastError:=5; LastErrorMsg:=TraduitGa(TexteMsgParboni[LastError]);
    SetFocusControl('APB_PRORATA');
    Exit;
  End;


  if (zpro.checked) and (GetField('APB_PRIPRO') = 0) then
  Begin
    LastError:=6; LastErrorMsg:=TraduitGa(TexteMsgParboni[LastError]);
    SetFocusControl('APB_PRIPRO');
    Exit;
  End;
  if not(zpro.checked) and (GetField('APB_PRIPRO') <> 0) then
  Begin
    LastError:=7; LastErrorMsg:=TraduitGa(TexteMsgParboni[LastError]);
    SetFocusControl('APB_PRIPRO');
    Exit;
  End;
  if (GetField('APB_RESSOURCE') <> '') and (GetField('APB_PRIRESS') = 0) then
  Begin
    LastError:=6; LastErrorMsg:=TraduitGa(TexteMsgParboni[LastError]);
    SetFocusControl('APB_PRIRESS');
    Exit;
  End;
  if (GetField('APB_RESSOURCE') = '') and (GetField('APB_PRIRESS') <> 0) then
  Begin
    LastError:=7; LastErrorMsg:=TraduitGa(TexteMsgParboni[LastError]);
    SetFocusControl('APB_PRIRESS');
    Exit;
  End;

  if (GetField('APB_RESSAFF') <> '') and (GetField('APB_PRIRESSAFF') = 0) then
  Begin
    LastError:=6; LastErrorMsg:=TraduitGa(TexteMsgParboni[LastError]);
    SetFocusControl('APB_PRIRESSAFF');
    Exit;
  End;
  if (GetField('APB_RESSAFF') = '') and (GetField('APB_PRIRESSAFF') <> 0) then
  Begin
    LastError:=7; LastErrorMsg:=TraduitGa(TexteMsgParboni[LastError]);
    SetFocusControl('APB_PRIRESSAFF');
    Exit;
  End;

   // Controle sur ordres de priorité
  za  := GetField('APB_PRIPRO');
  zb  := GetField('APB_PRIRESSAFF');
  zc  := GetField('APB_PRIRESS');

  if (za = zb) or (za=zc) or (zb=zc) then
   Begin
  	LastError:=2; LastErrorMsg:=TraduitGa(TexteMsgParboni[LastError]);
    SetFocusControl('APB_PRIPRO');
    Exit;
  End;
  if (za+zb+zc <> 3) and (za+zb+zc <> 6)  then
   Begin
  	LastError:=2; LastErrorMsg:=TraduitGa(TexteMsgParboni[LastError]);
    SetFocusControl('APB_PRIPRO');
    Exit;
  End;

     

end ;

procedure TOM_PARBONI.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_PARBONI.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_PARBONI.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_PARBONI.OnArgument ( S: String ) ;
var     ComboTypeArticle : THDBValComboBox;
begin
  Inherited ;
  //mcd 05/03/02
ComboTypeArticle:=THDbValComboBox(GetControl('APB_TYPEARTICLE'));
ComboTypeArticle.plus:=PlusTypeArticle;
end ;

procedure TOM_PARBONI.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_PARBONI.OnCancelRecord ;
begin
  Inherited ;
end ;

Procedure AFLanceFiche_ParamBoni;
begin
AGLLanceFiche ('AFF','AFPARBONI','','','');
end;

Initialization
  registerclasses ( [ TOM_PARBONI ] ) ; 
end.
