{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 11/07/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTSELLIGBORD ()
Mots clefs ... : TOF;BTSELLIGBORD
*****************************************************************}
Unit BTSELLIGBORD_TOF ;

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
     eMul, 
     uTob, 
  	 maineagl,
{$ENDIF}
		 Hqry,
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1,
     HMsgBox,
     HTB97,
     UTOB,
     AglInit,
     EntGc,
     UTOF ;

Type
  TOF_BTSELLIGBORD = Class (TOF)
  private
  	TheTOBResult : TOB;
//    BOuvrir : TToolbarButton97;
    procedure GSDblClick (Sender : TObject);
    procedure BouvrirClick (Sender : TOBject);
    procedure AddReference(QQ : THQuery);
		function PlusNaturePiece (XXWherePlus : string) : string;
 public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

function GetArticlesFromBordereaux (stWhere : string; Article : string; NaturePiece : string) : boolean;

Implementation

function GetArticlesFromBordereaux (stWhere : string; Article : string; NaturePiece : string) : boolean;
begin
	result := false;
  AGLLanceFiche ('BTP','BTSELLIGBORD','','','XXWHERE='+stWhere+';GL_REFARTTIERS='+article+';GL_NATUREPIECEG='+NaturePiece);
  if (theTOb <> nil) and (TheTob.detail.count > 0)  then result := true;
end;

{TOF_BTSELLIGBORD}

procedure TOF_BTSELLIGBORD.AddReference (QQ : THQUery);
var UneTOB : TOB;
    Nature,Souche : string;
    Indice,Numero,NumOrdre : integer;
    GetCedetail : string;
begin
  Nature := QQ.FindField('GL_NATUREPIECEG').AsString;
  Souche := QQ.FindField('GL_SOUCHE').AsString;
  Numero := QQ.FindField('GL_NUMERO').AsInteger;
  Indice := QQ.FindField('GL_INDICEG').AsInteger;
  NumOrdre := QQ.FindField('GL_NUMORDRE').AsInteger;
  GetCeDetail := QQ.FindField('GLC_GETCEDETAIL').AsString;
	UneTOB := TOB.Create ('LIGNE',TheTOBresult,-1);
  UneTOB.PutValue('GL_NATUREPIECEG',Nature);
  UneTOB.PutValue('GL_SOUCHE',Souche);
  UneTOB.PutValue('GL_NUMERO',Numero);
  UneTOB.PutValue('GL_INDICEG',Indice);
  UneTOB.PutValue('GL_NUMORDRE',Numordre);
  UneTOB.AddChampSupValeur ('GLC_GETCEDETAIL',GetCedetail);
end;

procedure TOF_BTSELLIGBORD.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTSELLIGBORD.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTSELLIGBORD.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTSELLIGBORD.OnLoad ;
begin
  Inherited ;
end ;

function TOF_BTSELLIGBORD.PlusNaturePiece (XXWherePlus : string) : string;
var lesChamps,UnChamps : string;
		Debut : boolean;
    LesNatures : string;
begin
	result := '';
	if XXWherePlus = '' then exit;
  LesChamps := XXWherePlus;
  Debut := true;
  LesNatures := '';
  repeat
  	UnChamps := READTOKENST (LesChamps);
    if Unchamps <> '' then
    begin
    	if debut then BEGIN LesNatures := '"'+UnChamps+'"'; debut := false; END
      				 else BEGIN LesNatures := LesNatures + ',"'+UnChamps+'"'; END;
    end;
  until UnChamps = '';
  result := ' AND (GL_TYPEARTICLE IN ('+lesNatures+'))';
end;

procedure TOF_BTSELLIGBORD.OnArgument (S : String ) ;
var critere,champMul,ValMul,XXWherePlus,XXWhere : string;
		x : integer;
begin
  Inherited ;
  repeat
    Critere := uppercase(Trim(ReadTokenSt(S)));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        ChampMul := copy(Critere, 1, x - 1);
        ValMul := copy(Critere, x + 1, length(Critere));
        if ChampMul = 'XXWHERE' then
        begin
          XXWhere := ValMul;
        end else
        if ChampMul = 'GL_REFARTTIERS' then
        begin
        	THEDIt(GEtControl('GL_REFARTTIERS')).Text := Valmul;
        end else
        if ChampMul = 'GL_NATUREPIECEG' then
        begin
        	XXWherePlus := GetInfoParPiece(ValMul, 'GPP_TYPEARTICLE');
        end else
        ;
      end;
    end;
  until Critere = '';
  if XXWherePlus <> '' then
  Begin
  	XXWhere := XXWhere + PlusNaturePiece (XXWherePlus) + ' AND (GL_REFARTTIERS<>"")';
  end;
  THEdit(GetControl('XX_WHERE')).text := XXWhere;
  TheTOBResult := TOB.Create('LE RESULTAT',nil,-1);
  TToolbarButton97 (GetControl('BOUVRIR')).OnClick := BouvrirClick;
  TFMul(GetControl('Fliste')).OnDblClick := GSDblClick;
end ;

procedure TOF_BTSELLIGBORD.OnClose ;
begin
  Inherited ;
  if TheTOBresult.detail.count = 0 then BEGIN TheTOBResult.free; TheTOBResult:= Nil END;
  TheTOB := TheTOBresult;
end ;

procedure TOF_BTSELLIGBORD.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTSELLIGBORD.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTSELLIGBORD.BouvrirClick(Sender: TOBject);
var iInd : integer;
//    TobRetour,TobRF : TOB;
begin
  with TFMul(Ecran) do
  begin
    if not FListe.AllSelected then
    begin
      if FListe.NbSelected = 0 then
      begin
        AddReference (Q);
      end else
      begin
        for iInd := 0 to FListe.NbSelected -1 do
        begin
          FListe.GotoLeBookMark(iInd);
  {$IFDEF EAGLCLIENT}
          Q.TQ.Seek (FListe.Row-1) ;
  {$ENDIF}
        AddReference (Q);
        end;
      end;
    end else
    begin
      Q.First;
      While not Q.Eof do
      begin
        AddReference (Q);
        Q.Next;
      end;
    end;
  end;
  ecran.ModalResult := MrOk;
end;

procedure TOF_BTSELLIGBORD.GSDblClick(Sender: TObject);
begin
  BouvrirClick (self);
end;

Initialization
  registerclasses ( [ TOF_BTSELLIGBORD ] ) ;
end.

