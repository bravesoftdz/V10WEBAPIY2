{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 24/01/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTOUVRAGE_MUL ()
Mots clefs ... : TOF;BTOUVRAGE_MUL
*****************************************************************}
Unit BTOUVRAGE_MUL_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     Fe_Main,
     HDB,EdtREtat,
{$else}
     MaineAGL,
     eMul,MaineAGL,
{$ENDIF}
     uTob,
		 Aglinit,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     Ent1,
     HMsgBox,
     HTB97,
     Menus,
     UTOF,
     BTMUL_TOF,
     NomenUtil,UentCommun ;

Type
  TOF_BTOUVRAGE_MUL = Class (TOF_BTMUL)
  private
  	ButInsert : TToolbarButton97;
    BDuplicat : TToolbarButton97;
    ffirst : boolean;
{$IFDEF EAGLCLIENT}
       GS : THGrid;
{$ELSE}
       GS : THDBGrid;
{$ENDIF}
  	procedure BimprimerClick (Sender : TObject);
    procedure LanceEdition (TheTOBParam : TOB);
    procedure AddChampEdt(TOBOUV: TOB);
    procedure AddChampSupPEdt(TheTOBParam: TOB);
    procedure SetChampDetail(TOBOUV, TheTOBParam, TOBL: TOB);
    procedure SetChampEntete(TOBOUV, TheTOBParam, TOBL: TOB);
    procedure GereEditionOuvrage(CodeArticle: string; TOBEDT,TheTOBParam: TOB);
    procedure SetChampFinEntete(TOBOUV, TheTOBParam, TOBL: TOB);
    procedure SetChampCplEntete(TOBOUV, TheTOBParam, TOBL: TOB);
    procedure RemplacecolonnesFamilles;
    procedure Duplicationclick(Sender: Tobject);


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

Implementation
uses paramsoc;

procedure TOF_BTOUVRAGE_MUL.AddChampSupPEdt (TheTOBParam : TOB);
begin
  TheTOBParam.AddChampSupValeur ('VALPA','-');
  TheTOBParam.AddChampSupValeur ('VALPR','-');
  TheTOBParam.AddChampSupValeur ('VALPV','-');
  TheTOBParam.AddChampSupValeur ('OPTSOUSDET','-');
  TheTOBParam.AddChampSupValeur ('SCODE','-');
  TheTOBParam.AddChampSupValeur ('SLIBELLE','-');
  TheTOBParam.AddChampSupValeur ('SQTE','-');
  TheTOBParam.AddChampSupValeur ('SUNITE','-');
  TheTOBParam.AddChampSupValeur ('OK','-');
end;
procedure TOF_BTOUVRAGE_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTOUVRAGE_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTOUVRAGE_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTOUVRAGE_MUL.OnLoad ;
begin
  Inherited ;
  if GetparamSocsecur('SO_BTRECHARTAUTO','X') then
  begin
		ffirst := false;
  end;
  if ffirst then ffirst := false
  					else THEdit(GetControl('XX_WHERE2')).Text := ''; 
  RemplacecolonnesFamilles;
end ;

procedure TOF_BTOUVRAGE_MUL.OnArgument (S : String ) ;
Var
    Critere	: string;
    Champ		: string;
    Valeur	: string;
    i_ind 	: integer;
    MenuPop : Tpopupmenu;
    CC      : THValComboBox;
begin
  Inherited ;
  ffirst := true;
	Critere:=(Trim(ReadTokenSt(S)));

  While (Critere <> '') do
  BEGIN
    i_ind:=pos(':',Critere);
    if i_ind = 0 then i_ind:=pos('=',Critere);
    if i_ind <> 0 then
    begin
       Champ:=copy(Critere,1,i_ind-1);
       Valeur:=Copy (Critere,i_ind+1,length(Critere)-i_ind);
    end else
    begin
       Champ := Critere;
    end;
    Critere:=(Trim(ReadTokenSt(S)));
  END;
  ButInsert := TToolbarButton97 (Getcontrol('Binsert'));
  MenuPop := TpopupMenu(GetControl ('POPCREATOUV'));
  ButInsert.DropdownMenu := MenuPop;
  ButInsert.Width := 35;
  ButInsert.DropDownAlways := true;

  //
  BDuplicat := TtoolBarButton97 (Getcontrol('B_DUPLICATION'));
  BDuplicat.OnClick := DuplicationClick;
  //

  TToolBarButton97(GetControl('BImprimer')).onclick := BimprimerClick;
  GS := THDbGrid (GetControl('Fliste'));

  //uniquement en line
  //SetControlVisible('PCOMPLEMENT',false);
  //TFMUL(Ecran).SetDbliste('BTMULOUVRAGE_S1');

  //Gestion Restriction Domaine
  CC:=THValComboBox(GetControl('GA_DOMAINE')) ;
  if CC<>Nil then PositionneDomaineUser(CC) ;

end ;

procedure TOF_BTOUVRAGE_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTOUVRAGE_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTOUVRAGE_MUL.OnCancel () ;
begin
  Inherited ;
end ;

Procedure TOF_BTOUVRAGE_MUL.Duplicationclick(Sender : Tobject);
Var CodeArticle : String;
begin

  CodeArticle := GS.datasource.dataset.FindField('GA_ARTICLE').AsString;

  if GetcontrolText('GA_TYPEARTICLE') <> 'RES' then
    V_PGI.DispatchTT(7,taCreat, CodeArticle,'','DUPLICATION=' + CodeArticle + ';TYPEARTICLE='+ GetcontrolText('GA_TYPEARTICLE'))
  else
    V_PGI.DispatchTT(6,taCreat, CodeArticle,'','DUPLICATION=' + CodeArticle);

end;

procedure TOF_BTOUVRAGE_MUL.BimprimerClick(Sender: TObject);
var TheTOBParam : TOB;
begin
  if (GS.nbSelected = 0) and (not GS.AllSelected) then
  begin
  	PgiInfo (Traduirememoire('Vous n''avez aucune sélection active'),ecran.caption) ;
  	exit;
  end;
  TheTOBParam := TOB.Create ('PARAMETRES D EDITION',nil,-1);
  AddChampSupPEdt (TheTOBParam);
  TheTOB := TheTOBParam;
  AGLLanceFiche ('BTP','BTOUVRAGE_PEDT','','','');
  if (TheTOB <> nil) and (TheTOB.GetValue('OK')='X') then
  begin
     LanceEdition (TheTOBParam);
  end;
  TheTOB := nil;
  FreeAndNil(TheTOBParam);
  GS.ClearSelected ;
end;

procedure TOF_BTOUVRAGE_MUL.LanceEdition(TheTOBParam: TOB);
var Q : TQuery;
		TOBEDT : TOB;
    i : integer;
    CodeArticle : string;
begin
	TOBEDT := TOB.Create ('LES LIGNES DEDITION',nil,-1);
	if (GS.AllSelected) then
  begin
    Q:=TFMul(ecran).Q;
    Q.First;
    while Not Q.EOF do
    begin
    	CodeArticle:=Q.FindField('GA_ARTICLE').AsString;
      GereEditionOuvrage(CodeArticle,TOBEDT,TheTOBParam);
      Q.next;
    end;
  end else
  begin
    for i:=0 to GS.nbSelected-1 do
    begin
      GS.GotoLeBookmark(i);
      CodeArticle:=GS.datasource.dataset.FindField('GA_ARTICLE').AsString;
      GereEditionOuvrage(CodeArticle,TOBEDT,TheTOBParam);
    end;
  end;

  if TOBEDT.detail.count > 0 then
  begin
  	LanceEtatTOB ('E','ART','BOU',TOBEDT,True,false,false,nil,'','',false);

  end;
  TOBEDT.free;
end;


procedure TOF_BTOUVRAGE_MUL.GereEditionOuvrage(CodeArticle : string; TOBEDT,TheTOBParam : TOB);
var QQ : TQuery;
		TOBLIGNE,TOBIntENT,TOBL,TOBDOUV,TOBIntLIG,TOBE : TOB;
    Indice,IndiceD: integer;
begin
  TOBIntENT := TOB.Create ('LES OUVRAGES',nil,-1);
  TOBIntLIG := TOB.Create ('LES OUVRAGES',nil,-1);

  QQ := OpenSql ('SELECT NOMENENT.*,GA_QUALIFUNITEVTE FROM NOMENENT LEFT JOIN ARTICLE ON GA_ARTICLE=GNE_ARTICLE '+
                 'WHERE GNE_ARTICLE="'+CodeArticle+'"',True);
  TOBIntENT.loadDetailDb('NOMENENT','','',QQ,False);
  ferme (QQ);
  for Indice := 0 to TOBIntENT.detail.count -1 do
  begin
    TOBE := TOBIntENT.detail[Indice];
    TOBLIGNE := TOB.Create ('UNE LIGNE',TOBEDT,-1);
    AddChampEdt (TOBLIGNE);
    SetChampEntete(TOBLIGNE,TheTOBParam,TOBE);
    if TheTOBParam.GetValue('OPTSOUSDET')='X' then
    begin
      TOBLIGNE := TOB.Create ('UNE LIGNE',TOBEDT,-1);
      AddChampEdt (TOBLIGNE);
    	SetChampCplEntete(TOBLIGNE,TheTOBParam,TOBE);
      //
      TOBIntLIG.clearDetail;
      QQ := OpenSql ('SELECT NOMENLIG.*,ARTICLE.GA_QUALIFUNITEVTE FROM NOMENLIG LEFT JOIN ARTICLE ON GA_ARTICLE=GNL_ARTICLE '+
                     'WHERE GNL_NOMENCLATURE="'+TOBE.GetValue('GNE_NOMENCLATURE')+'"',True);
      TOBIntLIG.loadDetailDb('NOMENLIG','','',QQ,False);
      ferme (QQ);
      for IndiceD := 0 to TOBIntLIG.detail.count -1 do
      begin
        TOBL := TOBIntLIG.detail[IndiceD];
        TOBLIGNE := TOB.Create ('UNE LIGNE',TOBEDT,-1);
        AddChampEdt (TOBLIGNE);
        SetChampDetail(TOBLIGNE,TheTOBParam,TOBL);
      end;
      TOBIntLIG.clearDetail;
    end;
    TOBLIGNE := TOB.Create ('UNE LIGNE',TOBEDT,-1);
    AddChampEdt (TOBLIGNE);
    SetChampFinEntete(TOBLIGNE,TheTOBParam,TOBL);
  end;
  TOBIntENT.free;
  TOBIntLIG.free;
end;

procedure TOF_BTOUVRAGE_MUL.AddChampEdt(TOBOUV: TOB);
begin
  TOBOUV.AddChampSupValeur('TYPEL',''); // Type de ligne (ENT Entete et LIG Sous detail)
  TOBOUV.AddChampSupValeur('VALPA','-');
  TOBOUV.AddChampSupValeur('VALPR','-');
  TOBOUV.AddChampSupValeur('VALPV','-');
  TOBOUV.AddChampSupValeur('MONTANTPA',0);
  TOBOUV.AddChampSupValeur('MONTANTPR',0);
  TOBOUV.AddChampSupValeur('MONTANTPV',0);
  TOBOUV.AddChampSupValeur('GNE_NOMENCLATURE','');
  TOBOUV.AddChampSupValeur('CODE','');
  TOBOUV.AddChampSupValeur('LIBELLE','');
  TOBOUV.AddChampSupValeur('QTE',0);
  TOBOUV.AddChampSupValeur('UNITE','');
  TOBOUV.AddChampSupValeur('QTEDUPV',0);
  TOBOUV.AddChampSupValeur ('SCODE','-');
  TOBOUV.AddChampSupValeur ('SLIBELLE','-');
  TOBOUV.AddChampSupValeur ('SQTE','-');
  TOBOUV.AddChampSupValeur ('SUNITE','-');
end;

procedure TOF_BTOUVRAGE_MUL.SetChampFinEntete(TOBOUV,TheTOBParam,TOBL : TOB);
begin
  TOBOUV.PutValue('TYPEL','FIN');
end;

procedure TOF_BTOUVRAGE_MUL.SetChampCplEntete (TOBOUV,TheTOBParam,TOBL : TOB);
begin
  TOBOUV.PutValue('TYPEL','CPL');
  TOBOUV.PutValue('QTEDUPV',TOBL.GetValue('GNE_QTEDUDETAIL'));
  TOBOUV.PutValue('UNITE',TOBL.GetValue('GA_QUALIFUNITEVTE'));
end;

procedure TOF_BTOUVRAGE_MUL.SetChampEntete(TOBOUV,TheTOBParam,TOBL : TOB);
var Valeurs : T_Valeurs;
		ValPa,ValPr,ValPv : string;
begin
	ValPa := TheTOBParam.GetValue('VALPA');
	ValPR := TheTOBParam.GetValue('VALPR');
	ValPV := TheTOBParam.GetValue('VALPV');
  TOBOUV.PutValue('TYPEL','ENT');
  TOBOUV.PutValue('GNE_NOMENCLATURE',TOBL.GetValue('GNE_NOMENCLATURE'));
  TOBOUV.PutValue('CODE',Copy(TOBL.GetValue('GNE_NOMENCLATURE'),1,18));
  TOBOUV.PutValue('LIBELLE',TOBL.GetValue('GNE_LIBELLE'));
  TOBOUV.PutValue('QTEDUPV',TOBL.GetValue('GNE_QTEDUDETAIL'));
  TOBOUV.PutValue('UNITE',TOBL.GetValue('GA_QUALIFUNITEVTE'));
  TOBOUV.PutValue('VALPA',ValPa);
  TOBOUV.PutValue('VALPR',ValPR);
  TOBOUV.PutValue('VALPV',VALPV);
  if (ValPa ='X') or (VAlPR ='X') or (ValPV ='X') then
  begin
    ValoriseOuvrage ( TOBL.getvalue('GNE_NOMENCLATURE'),nil,nil,valeurs);
    if ValPV='X' then TOBOUV.putvalue('MONTANTPV',Valeurs[2]);
    if ValPA='X' then TOBOUV.putvalue('MONTANTPA',Valeurs[0]);
    if ValPR='X' then TOBOUV.putvalue('MONTANTPR',Valeurs[1]);
  end;

end;

procedure TOF_BTOUVRAGE_MUL.SetChampDetail(TOBOUV,TheTOBParam,TOBL : TOB);
var Valeurs : T_Valeurs;
		ValPa,ValPr,ValPv : string;
begin
	ValPa := TheTOBParam.GetValue('VALPA');
	ValPR := TheTOBParam.GetValue('VALPR');
	ValPV := TheTOBParam.GetValue('VALPV');
  TOBOUV.PutValue('TYPEL','LIG');
  TOBOUV.PutValue('GNE_NOMENCLATURE',TOBL.GetValue('GNL_NOMENCLATURE'));
  TOBOUV.PutValue('VALPA',ValPa);
  TOBOUV.PutValue('VALPR',ValPR);
  TOBOUV.PutValue('VALPV',VALPV);
  TOBOUV.PutValue ('SCODE',TheTOBParam.GetValue('SCODE'));
  TOBOUV.PutValue ('SLIBELLE',TheTOBParam.GetValue('SLIBELLE'));
  TOBOUV.PutValue ('SQTE',TheTOBParam.GetValue('SQTE'));
  TOBOUV.PutValue ('SUNITE',TheTOBParam.GetValue('SUNITE'));
  if TOBOUV.GetValue('SCODE')='X' then TOBOUV.PutValue('CODE',Copy(TOBL.GetValue('GNL_CODEARTICLE'),1,18));
  if TOBOUV.GetValue('SLIBELLE')='X' then TOBOUV.PutValue('LIBELLE',TOBL.GetValue('GNL_LIBELLE'));
  if TOBOUV.GetValue('SQTE')='X' then TOBOUV.PutValue('QTE',TOBL.GetValue('GNL_QTE'));
  if TOBOUV.GetValue('SUNITE')='X' then TOBOUV.PutValue('UNITE',TOBL.GetValue('GA_QUALIFUNITEVTE'));
end;

procedure TOF_BTOUVRAGE_MUL.RemplacecolonnesFamilles;
var i : integer;
		Gr : THDbgrid;
    stChamp,Libelle : string;
begin
	Gr := TFMul(ecran).fliste;
	For i:=0 to Gr.Columns.Count-1 do
  Begin
    StChamp := TFMul(Ecran).Q.FormuleQ.GetFormule(Gr.Columns[i].FieldName);
        if copy(UpperCase (stChamp),1,6)='LIBPRE' then
    begin
      libelle := RechDom('GCLIBFAMILLE','LF'+Copy(stChamp,7,1),false);
{$IFNDEF AGL581153}
			TFMul(ecran).SetDisplayLabel (StChamp,TraduireMemoire(Libelle));
{$else}
			TFMul(ecran).SetDisplayLabel (i,TraduireMemoire(Libelle));
{$endif}
    end else
    if copy(UpperCase (stChamp),1,6)='LIBOUV' then
    begin
      libelle := RechDom('BTLIBOUVRAGE','BO'+Copy(stChamp,7,1),false);
{$IFNDEF AGL581153}
			TFMul(ecran).SetDisplayLabel (StChamp,TraduireMemoire(Libelle));
{$else}
			TFMul(ecran).SetDisplayLabel (i,TraduireMemoire(Libelle));
{$endif}
    end;
  end;
end;

Initialization
  registerclasses ( [ TOF_BTOUVRAGE_MUL ] ) ;
end.

