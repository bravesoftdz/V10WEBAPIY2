{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 20/02/2010
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BARTVERSIONNABLE ()
Mots clefs ... : TOF;BARTVERSIONNABLE
*****************************************************************}
Unit BARTVERSIONNABLE_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,HDB,fe_main,
{$else}
     eMul,Main_eagl,
{$ENDIF}
     uTob,
     forms,
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1,
     menus,
     HMsgBox,
     AglInit,
     UTOF ; 

Type
  TOF_BARTVERSIONNABLE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	POPGS : TPopupMenu;
    AffectVers : TMenuItem;
  	procedure RemplacecolonnesFamilles;
    procedure Affecteversion (Sender : Tobject);

  end ;

Implementation

procedure TOF_BARTVERSIONNABLE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BARTVERSIONNABLE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BARTVERSIONNABLE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BARTVERSIONNABLE.OnLoad ;
begin
  Inherited ;
  RemplacecolonnesFamilles;
end ;

procedure TOF_BARTVERSIONNABLE.OnArgument (S : String ) ;
begin
  Inherited ;
  POPGS := TPopupMenu(GetControl('POPGS'));
  AffectVers := TmenuItem(getControl('AffectVers'));
  AffectVers.onclick := Affecteversion;
end ;

procedure TOF_BARTVERSIONNABLE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BARTVERSIONNABLE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BARTVERSIONNABLE.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BARTVERSIONNABLE.RemplacecolonnesFamilles;
var i : integer;
		Gr : THDbgrid;
    stChamp,Libelle : string;
begin
	Gr := TFMul(ecran).fliste;
	For i:=0 to Gr.Columns.Count-1 do
  Begin
    StChamp := TFMul(Ecran).Q.FormuleQ.GetFormule(Gr.Columns[i].FieldName);
    if copy(UpperCase (stChamp),1,7)='LIBPARC' then
    begin
      libelle := RechDom('BTLIBARTPARC','PA'+Copy(stChamp,8,1),false);
{$IFNDEF AGL581153}
			TFMul(ecran).SetDisplayLabel (StChamp,TraduireMemoire(Libelle));
{$else}
			TFMul(ecran).SetDisplayLabel (i,TraduireMemoire(Libelle));
{$endif}
    end;
  end;
end;

procedure TOF_BARTVERSIONNABLE.Affecteversion(Sender: Tobject);
VAR i : integer;
		CodeArticle,Articles : string;
    first : boolean;
    TOBTRA : TOB;
    Q : Tquery;
begin
	Articles := '';
  first := true;
  if TFMul(ecran).Fliste.AllSelected then
  BEGIN
    Q:=TFMul(ecran).Q;
    Q.First;
    while Not Q.EOF do
    BEGIN
      CodeArticle:=Q.FindField('GA_CODEARTICLE').AsString;
      if first then first := false else Articles := Articles+';';
      Articles := Articles+CodeArticle;
      Q.NEXT;
    END;
    TFMul(ecran).Fliste.AllSelected:=False;
  END ELSE
  BEGIN
    for i:=0 to TFMul(ecran).Fliste.nbSelected-1 do
    begin
      TFMul(ecran).Fliste.GotoLeBookmark(i);
      CodeArticle:=TFMul(ecran).Fliste.datasource.dataset.FindField('GA_CODEARTICLE').AsString;
      if first then first := false else Articles := Articles+';';
      Articles := Articles+CodeArticle;
    end;
	END;
  if Articles <> '' then
  begin
  	TOBTRa := TOB.Create ('LE PASSAGE',nil,-1);
    TOBTRA.AddChampSupValeur ('ARTICLES',Articles);
    TOBTRA.AddChampSupValeur ('FORCE','X');
    TOBTRA.AddChampSupValeur ('RETOUR','');
    TheTOB := TOBTra;
    AGLLanceFiche('BTP','BTCHANGEVERS','','','ACTION=MODIFICATION');
    TheTOB := nil;
    TOBTRA.free;
  end;
  TFMul(ecran).Fliste.ClearSelected;
end;

Initialization
  registerclasses ( [ TOF_BARTVERSIONNABLE ] ) ; 
end.
