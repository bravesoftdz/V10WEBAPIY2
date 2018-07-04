{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 20/08/2012
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTRECALCPMAPNOSTK ()
Mots clefs ... : TOF;BTRECALCPMAPNOSTK
*****************************************************************}
Unit BTRECALCPMAPNOSTK_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes,
     HTB97,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     FE_Main,
     HDB,
     dbgrids,
{$else}
     eMul,
     uTob,
     Maineagl,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     ParamSoc,
     uTOB,
     UTOF ;

Type
  TOF_BTRECALCPMAPNOSTK = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  Private
    DateDay     : THEdit;
    DateYear    : THEdit;
    XX_WHERE    : THEdit;
    TYPEARTICLE : THMultiValComboBox;
    TENUESTOCK  : THcheckBox;
    CHKFERME    : THCheckBox;
    Fliste      : THDBGrid;
    BOuvrir     : TToolBarButton97;
    //
    procedure GetObjects;
    procedure VisibleScreenObject;
    procedure SetScreenEvents;
    procedure DateYearOnExit(Sender: Tobject);
    procedure OnValideClick(Sender: Tobject);
    procedure TraitementCalculPMAP(TobArt : TOB);
    procedure MAJTableArticle(TobArt: TOB);
    procedure DatedayOnExit(Sender: Tobject);
  end ;

Implementation

procedure TOF_BTRECALCPMAPNOSTK.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTRECALCPMAPNOSTK.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTRECALCPMAPNOSTK.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTRECALCPMAPNOSTK.OnLoad ;
var StWhere : String;
begin
  Inherited ;

  //Chargement de la requete SQL
  //StSql   := 'select ga_article,ga_libelle from article where ga_Ferme='-' And ga_tenuestock="-" AND GA_TYPEARTICLE IN ("MAR","ARP") AND ';

  Stwhere := '(select count(*) from ligne where gl_article=ga_article ';
  Stwhere := StWhere + 'and gl_naturepieceg="FF" ';
  StWhere := Stwhere + 'and gl_datepiece>="' + USDATE(Dateyear) + '" ';
  Stwhere := Stwhere + 'and gl_datepiece <= "' + USDATE(DateDay) + '") > 0';

  XX_WHERE.Text := StWhere;

end ;

procedure TOF_BTRECALCPMAPNOSTK.OnArgument (S : String ) ;
begin
  Inherited ;

  GetObjects;

  SetScreenEvents;

  VisibleScreenObject;

  //initialisation des zones écran...
  if assigned(Dateday)  then DateDay.Text  := DateToStr(Now);
  if Assigned(Dateyear) then DateYear.Text := DateToStr(PlusDate(Now, -12, 'M'));

  XX_WHERE.Text       := '';
  TenueStock.Checked  := False;
  ChkFERME.Checked    := False;
  TypeArticle.Value   := 'ARP;MAR';

end ;

procedure TOF_BTRECALCPMAPNOSTK.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTRECALCPMAPNOSTK.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTRECALCPMAPNOSTK.OnCancel () ;
begin
  Inherited ;
end ;

Procedure TOF_BTRECALCPMAPNOSTK.GetObjects;
begin

  if assigned(Getcontrol('DATEPIECE'))      then DateDay      := THEdit(ecran.FindComponent('DATEPIECE'));
  if assigned(Getcontrol('DATEPIECE_'))     then DateYear     := THEdit(ecran.FindComponent('DATEPIECE_'));
  if Assigned(GetControl('XX_WHERE'))       then XX_WHERE     := THEdit(ecran.findComponent('XX_WHERE'));
  if Assigned(GetControl('GA_TYPEARTICLE')) then TYPEARTICLE  := THMultiValComboBox(ecran.findComponent('GA_TYPEARTICLE'));
  if Assigned(GetControl('GA_TENUESTOCK'))  then TENUESTOCK   := THCheckBox(ecran.findComponent('GA_TENUESTOCK'));
  if Assigned(GetControl('GA_FERME'))       then CHKFERME     := THCheckBox(ecran.findComponent('GA_FERME'));
  if Assigned(Getcontrol('FLISTE'))         then Fliste       := THDbGrid (ecran.findComponent('FLISTE'));
  if Assigned(Getcontrol('BOuvrir'))        then Bouvrir      := TToolBarButton97(ecran.findComponent('BOuvrir'));

end;

procedure TOF_BTRECALCPMAPNOSTK.VisibleScreenObject;
begin

  TypeArticle.visible := False;
  TenueStock.visible  := False;
  CHKFERME.Visible    := False;
  XX_Where.Visible    := False;

end;

procedure TOF_BTRECALCPMAPNOSTK.SetScreenEvents;
begin

  if assigned(Dateyear)  then DateYear.OnExit := DateYearOnExit;
  if assigned(Dateday)   then DateDay.OnExit  := DateDayOnExit;
  if Assigned(BOuvrir)   then Bouvrir.OnClick := OnValideClick;

end;

procedure TOF_BTRECALCPMAPNOSTK.DateyearOnExit(Sender : Tobject);
begin

  DateDay.Text := DateToStr(PlusDate(StrToDate(DateYear.text), 12, 'M'));

end;

procedure TOF_BTRECALCPMAPNOSTK.DatedayOnExit(Sender : Tobject);
begin

  DateYear.Text := DateToStr(PlusDate(StrToDate(Dateday.text), -12, 'M'));

end;

Procedure TOF_BTRECALCPMAPNOSTK.OnValideClick(Sender : Tobject);
var Article : String;
    i       : Integer;
    QQ      : TQuery;
    TOBArt  : Tob;
    TobTrait: Tob;
Begin

  if(FListe.NbSelected=0) and (not FListe.AllSelected) then
  begin
	   PGIInfo('Aucun élément sélectionné','');
  	 exit;
  end;

	if PGIAsk('Confirmez-vous le traitement ?','') <> mrYes then exit ;

  SourisSablier;

  TOBArt := tob.create('LES ARTICLES', nil, -1);

  try
    if FListe.AllSelected then
  	begin
    	QQ:=TFMul(Ecran).Q;
    	QQ.First;
	    while Not QQ.EOF do
      Begin
     	  Article := QQ.FindField('GA_ARTICLE').AsString;
        TOBTrait := TOB.Create ('CALCPMAP',TobArt,-1);
        TOBTrait.AddChampSupValeur('GA_ARTICLE', Article);
        TOBTrait.AddChampSupValeur('GA_PMAP', 0);
      	QQ.Next;
   	 end;
     TraitementCalculPMAP(TobArt);
    end
   	else
    begin
      for i:=0 to FListe.NbSelected-1 do
      begin
	      FListe.GotoLeBookmark(i);
        Article := Fliste.datasource.dataset.FindField('GA_ARTICLE').AsString;
        TOBTrait := TOB.Create ('CALCPMAP',TobArt,-1);
        TOBTrait.AddChampSupValeur('GA_ARTICLE',Article);
        TOBTrait.AddChampSupValeur('GA_PMAP', 0);
      end;
    end;
    TraitementCalculPMAP(TobArt);
    MAJTableArticle(TOBArt);
    if FListe.AllSelected then
      FListe.AllSelected:=False
    else
      FListe.ClearSelected;
	finally
  	SourisNormale;
    FreeAndNil(TobArt);
    TFMul(Ecran).BChercheClick(ecran);
  end;
end;

Procedure TOF_BTRECALCPMAPNOSTK.TraitementCalculPMAP(TobArt : TOB);
Var Ind   : Integer;
    TobL  : TOB;
    StSql : String;
    QQ    : Tquery;
    QteTotal  : Double;
    MtHtDev   : Double;
    PMAP      : Double;
begin

  if TobArt.Detail.count = 0 then exit;

  for ind := 0 to TobArt.detail.count -1  do
  begin
    TobL := TobArt.detail[Ind];
    StSql := 'SELECT Sum(GL_QTEFACT) As QteTotal, Sum(GL_MONTANTHTDEV) As MtHtDev FROM LIGNE ';
    StSQl := StSQl + 'WHERE GL_ARTICLE="' + TOBL.GetString('GA_ARTICLE') + '" ';
    StSQl := StSQl + 'AND   GL_NATUREPIECEG="FF" ';
    StSQl := StSQl + 'AND   GL_DATEPIECE >="' + USDATE(Dateyear) + '" ';
    StSQl := StSQl + 'AND   GL_DATEPIECE <= "' + USDATE(DateDay) + '"';
    QQ := OpenSQL(StSQL, false);
    if Not QQ.eof then
    begin
      QteTotal := QQ.FindField('QteTotal').AsFloat;
      MtHtDev  := QQ.FindField('MtHtDev').AsFloat;
      PMAP     := 0;
      if QteTotal <> 0 then PMAP := Arrondi(MtHtDev / QteTotal, V_PGI.OkDecP);
      TOBL.AddchampSupValeur('QTETOTAL', QteTotal);
      TOBL.AddchampSupValeur('MTHTDEV', MtHtDev);
      TOBL.PutValue('GA_PMAP', PMAP);
    end;
    ferme(QQ);
  end;

end;

Procedure TOF_BTRECALCPMAPNOSTK.MAJTableArticle(TobArt : TOB);
var ind   : integer;
    Art   : String;
    PMAP  : double;
    StSQL : String;
begin

  if TobArt.Detail.count = 0 then exit;

    // RAZ PMAP de tous les articles avant MAJ : modif BRL 17/04/2013
  StSQL := 'update article set ga_pmap=0 where ga_typearticle in ("ARP","MAR") and ga_tenuestock="-" and ga_ferme="-"';
  ExecuteSQL(StSQL);

  for ind := 0 to TobArt.detail.count -1  do
  begin
    Art := TobArt.Detail[Ind].GetString('GA_ARTICLE');
    PMAP:= TobArt.Detail[Ind].GetDouble('GA_PMAP');
    StSQL := 'UPDATE ARTICLE SET GA_PMAP = ' + strfpoint(PMAP) + ' WHERE GA_ARTICLE="' + Art + '"';
    ExecuteSQL(StSQL);
  end;


end;

Initialization
  registerclasses ( [ TOF_BTRECALCPMAPNOSTK ] ) ;
end.

