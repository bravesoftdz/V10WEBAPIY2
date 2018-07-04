{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 02/04/2014
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTMOUVSTK ()
Mots clefs ... : TOF;BTMOUVSTK
*****************************************************************}
Unit UTOFBTMOUVSTK;

Interface

Uses StdCtrls,
     Controls,
     uEntCommun,
     Classes,
{$IFNDEF EAGLCLIENT}
     Mul,FichList,FE_Main,DBCtrls,DBGrids,HDB,db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}EdtREtat,
{$else}
     eMul,
{$ENDIF}
     Hqry,
     HTB97,
     uTob,
     forms,
     sysutils, 
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     AGLInit,
     UTOF ;

Type
  TOF_BTMOUVSTK = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    stwhere : string;
    fUid : string;
    procedure CodeArticleElipsisClick(Sender: TObject);
    procedure fListeDblClick (Sender : TOBject);
    procedure BeforeAffichage;
    procedure BchercheClick (Sender : TObject);
    function TraduitPourOuvrage (Where : string) : string;
    function QteSousDetail (Qte : Double ; cledoc : r_cledoc; N1,N2,N3,N4,N5 : Integer) : double;
    function TraduitPourLigne(Where: string): string;
    procedure ScreenKeYDown (Sender: TObject; var Key: Word; Shift: TShiftState);
  end ;

Implementation
uses Facture,StockUtil;

procedure TOF_BTMOUVSTK.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTMOUVSTK.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTMOUVSTK.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTMOUVSTK.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTMOUVSTK.OnArgument (S : String ) ;
begin
  Inherited ;
  THEdit(GetControl('BTM_ARTICLE')).OnElipsisClick := CodeArticleElipsisClick;
  THDBGrid (GetControl('fListe')).OnDblClick := fListeDblClick;
  Ecran.OnKeyDown := ScreenKeYDown;
  fUid := AglGetGuid;
  SetControlText('BTM_UID',fUid);
  TToolbarButton97(GetControl('BchercheS')).onclick := BchercheClick;
end ;

procedure TOF_BTMOUVSTK.OnClose ;
begin
  ExecuteSQL('DELETE FROM BTTMPMOUVSTK WHERE BTM_UID="'+fUid+'"');
  Inherited ;
end ;

procedure TOF_BTMOUVSTK.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTMOUVSTK.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTMOUVSTK.CodeArticleElipsisClick(Sender: TObject);
var SWhere,StChamps : string;
		Article : Thedit;
    TheArticle : string;
begin
  Article := THEdit(getControl('BTM_ARTICLE'));
	sWhere := 'AND ((GA_TYPEARTICLE="MAR") OR (GA_TYPEARTICLE="ARP"))';
  if Article.Text <> '' then StChamps := 'GA_CODEARTICLE=' + Trim(Copy(Article.text, 1, 18)) + ';XX_WHERE=' + sWhere
		 										else StChamps := 'XX_WHERE=' + sWhere;

	TheArticle  := AGLLanceFiche('BTP', 'BTARTICLE_RECH', '', '', StChamps+';RECHERCHEARTICLE;STATUTART=UNI,DIM');
  Article.text := TheArticle;
end;

procedure TOF_BTMOUVSTK.fListeDblClick(Sender: TOBject);
var NaturePiece,Souche : string;
		Numero,Indice : Integer;
    cledoc : String;
    fListe : THDBGrid;
begin
  fListe := THDBGrid(GetControl('fListe'));
	NaturePiece:=Fliste.datasource.dataset.FindField('BTM_NATUREPIECEG').AsString;
	Souche:=Fliste.datasource.dataset.FindField('BTM_SOUCHE').AsString;
	Numero:=Fliste.datasource.dataset.FindField('BTM_NUMERO').AsInteger;
	Indice:=Fliste.datasource.dataset.FindField('BTM_INDICEG').AsInteger;
  cledoc := Naturepiece+';;'+Souche+';'+InttoStr(Numero)+';'+InttoStr(Indice)+';';
  AppelPiece([cledoc,'ACTION=CONSULTATION'],2);
end;

procedure TOF_BTMOUVSTK.BeforeAffichage;

function QteLigne (cledoc : R_CLEDOC) : Double;
var Q : TQuery;
begin
  Result := 1;
  Q := OpenSQL('SELECT GL_QTERESTE FROM LIGNE '+
              'WHERE '+
              'GL_NATUREPIECEG="'+Cledoc.NaturePiece+'" AND '+
              'GL_SOUCHE="'+Cledoc.souche+'" AND '+
              'GL_NUMERO='+IntToStr(Cledoc.NumeroPiece)+' AND '+
              'GL_INDICEG='+IntToStr(Cledoc.Indice)+' AND '+
              'GL_NUMLIGNE='+IntToStr(Cledoc.NumLigne), true,1,'',true);
  if not Q.Eof then result := Q.fields[0].AsFloat;
  ferme (Q);
end;

procedure SupprimeLivFromFourn;
var QQ : TQuery;
begin
  QQ := OpenSql ('SELECT * FROM BTTMPMOUVSTK WHERE BTM_UID="'+fUid+'" AND BTM_NATUREPIECEG IN ("LBT","BLC")',false);
  if not QQ.eof then
  begin
    QQ.First;
    While not QQ.eof do
    begin
      if not IsLivChantier (QQ.FindField('BTM_PIECEPRECEDENTE').AsString,QQ.FindField('BTM_PIECEORIGINE').AsString) then QQ.next
                                                                       else QQ.Delete;
    end;
  end;
  ferme (QQ);
end;


var SQl : string;
    QQ : TQuery;
    QteReelle : Double;
    Cledoc : R_CLEDOC;
    N1,N2,N3,N4,N5 : Integer;
begin
  ExecuteSQL('DELETE FROM BTTMPMOUVSTK WHERE BTM_UID="'+fUid+'"');
  // Chargement des lignes
  SQl := 'INSERT INTO BTTMPMOUVSTK '+
        '(BTM_NATUREDATA,BTM_UID, BTM_DATECREATION, BTM_NATUREPIECEG, BTM_SOUCHE, BTM_NUMERO, BTM_INDICEG, BTM_TIERS, '+
        'BTM_CODEARTICLE, BTM_ARTICLE, BTM_LIBELLE, BTM_DATEPIECE, BTM_QTESTOCK, BTM_PU, BTM_QUALIFMVT, '+
        'BTM_DEPOT, BTM_SENSPIECE, BTM_QUALIFUNITE,BTM_NUMLIGNE,BTM_N1,BTM_N2,BTM_N3,BTM_N4,BTM_N5,BTM_PIECEORIGINE,BTM_PIECEPRECEDENTE) '+
        ' SELECT "LIG","'+fUid+'",GL_DATECREATION,GL_NATUREPIECEG,GL_SOUCHE,GL_NUMERO,GL_INDICEG,GL_TIERS,'+
        'GL_CODEARTICLE,GL_ARTICLE,GL_LIBELLE,GL_DATEPIECE,GL_QTEFACT,PU,GL_QUALIFMVT,GL_DEPOT,GL_SENSPIECE,UNITE,GL_NUMLIGNE,0,0,0,0,0,GL_PIECEORIGINE,GL_PIECEPRECEDENTE '+
        'FROM BTLISTMSTOCK '+TraduitPourLigne(stWhere)+' AND GL_QTERESTE <> 0';
  //
  ExecuteSql(Sql);
  // Chargement des Détail d'ouvrages
  SQl := 'INSERT INTO BTTMPMOUVSTK '+
        '(BTM_NATUREDATA,BTM_UID, BTM_DATECREATION, BTM_NATUREPIECEG, BTM_SOUCHE, BTM_NUMERO, BTM_INDICEG, BTM_TIERS, '+
        'BTM_CODEARTICLE, BTM_ARTICLE, BTM_LIBELLE, BTM_DATEPIECE, BTM_QTESTOCK, BTM_PU, BTM_QUALIFMVT, '+
        'BTM_DEPOT, BTM_SENSPIECE, BTM_QUALIFUNITE,BTM_NUMLIGNE,BTM_N1,BTM_N2,BTM_N3,BTM_N4,BTM_N5,BTM_PIECEORIGINE,BTM_PIECEPRECEDENTE) '+
        ' SELECT "OUV","'+fUid+'",BLO_DATECREATION,BLO_NATUREPIECEG,BLO_SOUCHE,BLO_NUMERO,BLO_INDICEG,BLO_TIERS,'+
        'BLO_CODEARTICLE,BLO_ARTICLE,BLO_LIBELLE,BLO_DATEPIECE,BLO_QTEFACT,PU,BLO_QUALIFMVT,BLO_DEPOT,BLO_SENSPIECE,UNITE,BLO_NUMLIGNE,BLO_N1,BLO_N2,BLO_N3,BLO_N4,BLO_N5,GL_PIECEORIGINE,GL_PIECEPRECEDENTE '+
        'FROM BTLISTMSTOCKO '+TraduitPourOuvrage(stWhere);
  //
  ExecuteSql(Sql);

  SupprimeLivFromFourn;
  
  // Parcours des détails d'ouvrages
  QQ := OpenSql ('SELECT * FROM BTTMPMOUVSTK WHERE BTM_UID="'+fUid+'" AND BTM_NATUREDATA="OUV"',false);
  if not QQ.eof then
  begin
    QQ.First;
    While not QQ.eof do
    begin
      //
      Cledoc.NaturePiece := QQ.findfield('BTM_NATUREPIECEG').AsString;
      Cledoc.Souche := QQ.findfield('BTM_SOUCHE').AsString ;
      Cledoc.NumeroPiece := QQ.findfield('BTM_NUMERO').AsInteger ;
      Cledoc.Indice := QQ.findfield('BTM_INDICEG').AsInteger ;
      Cledoc.NumLigne  := QQ.findfield('BTM_NUMLIGNE').AsInteger ;
      //
      N1 :=QQ.findfield('BTM_N1').AsInteger;
      N2 :=QQ.findfield('BTM_N2').AsInteger;
      N3 :=QQ.findfield('BTM_N3').AsInteger;
      N4 :=QQ.findfield('BTM_N4').AsInteger;
      N5 :=QQ.findfield('BTM_N5').AsInteger;
      //
      QteReelle := QteLigne (Cledoc) * QteSousDetail( QQ.findfield('BTM_QTESTOCK').Asfloat, Cledoc,N1,N2,N3,N4,N5);
      //
      QQ.Edit;
      QQ.FindField('BTM_QTESTOCK').AsFloat := QteReelle;
      QQ.UpdateRecord;
      QQ.Next;
    end;
  end;
  ferme (QQ);
end;

procedure TOF_BTMOUVSTK.BchercheClick(Sender: TObject);
begin
  SetControlText('BTM_UID','');
  stWhere := RecupWhereCritere(TPageControl(GetControl('Pages'))) ;
  SetControlText('BTM_UID',fUid);
  BeforeAffichage;
  TToolbarButton97(GetControl('BCherche')).Click;
end;

function TOF_BTMOUVSTK.TraduitPourOuvrage(Where: string): string;
begin
  Result := FindEtReplace(Where,'BTM_','BLO_',true)
end;

function TOF_BTMOUVSTK.TraduitPourLigne(Where: string): string;
begin
  Result := FindEtReplace(Where,'BTM_','GL_',true)
end;

function TOF_BTMOUVSTK.QteSousDetail(Qte : Double ; cledoc: r_cledoc; N1, N2, N3, N4, N5: Integer): double;


  function ConstitueWhereFather(N1,N2,N3,N4,N5 : integer) : string;
  var level : Integer;
  begin
    Result := '';
    // Recherche su niveau actuel
    For level := 5 downto 1 do
    begin
      if Level = 5 then
      begin
        if N5 <> 0 then break;
      end else if Level = 4 then
      begin
        if N4 <> 0 then break;
      end else if Level = 3 then
      begin
        if N3 <> 0 then break;
      end else if Level = 2 then
      begin
        if N2 <> 0 then break;
      end else if Level = 1 then exit;
    end;

    if level > 1 then
    begin
      if level = 5 then
      begin
        Result := ' AND BLO_N5=0'+
                  ' AND BLO_N4='+InttoStr(N4)+
                  ' AND BLO_N3='+InttoStr(N3)+
                  ' AND BLO_N2='+InttoStr(N2)+
                  ' AND BLO_N1='+InttoStr(N1);
      end else if level = 4 then
      begin
        Result := ' AND BLO_N4=0 AND BLO_N3='+InttoStr(N3)+
                  ' AND BLO_N2='+InttoStr(N2)+
                  ' AND BLO_N1='+InttoStr(N1);
      end else if level = 3 then
      begin
        Result := ' AND BLO_N3=0 AND BLO_N2='+InttoStr(N2)+
                  ' AND BLO_N1='+InttoStr(N1);
      end else if level = 2 then
      begin
        Result := ' AND BLO_N2=0 AND BLO_N1='+InttoStr(N1);
      end;
    end;
  end;


var Q : TQuery;
    StSqlNivPere : string;
    NN1,NN2,NN3,NN4,NN5 : Integer;
    QteLoc : Double;
begin
  result := Qte;
  StSqlNivPere := ConstitueWhereFather(N1,N2,N3,N4,N5);
  if StSqlNivPere = '' then exit;
  Q := OpenSQL('SELECT BLO_NATUREPIECEG ,BLO_SOUCHE ,BLO_NUMERO ,'+
              'BLO_INDICEG ,BLO_QTEFACT,BLO_NUMLIGNE,'+
              'BLO_N1,BLO_N2,BLO_N3,BLO_N4,BLO_N5 '+
              'FROM LIGNEOUV '+
              'WHERE '+
              'BLO_NATUREPIECEG="'+Cledoc.NaturePiece+'" AND '+
              'BLO_SOUCHE="'+Cledoc.souche+'" AND '+
              'BLO_NUMERO='+IntToStr(Cledoc.NumeroPiece)+' AND '+
              'BLO_INDICEG='+IntToStr(Cledoc.Indice)+' AND '+
              'BLO_NUMLIGNE='+IntToStr(Cledoc.NumLigne)+
               StSqlNivPere, true,1,'',true);
  if Not Q.eof then
  begin
    //
    NN1 :=Q.findfield('BLO_N1').AsInteger;
    NN2 :=Q.findfield('BLO_N2').AsInteger;
    NN3 :=Q.findfield('BLO_N3').AsInteger;
    NN4 :=Q.findfield('BLO_N4').AsInteger;
    NN5 :=Q.findfield('BLO_N5').AsInteger;
    //
    QteLoc := Q.findField('BLO_QTEFACT').AsFloat;
    //
    Result := result * QteSousDetail(QteLoc,Cledoc,NN1,NN2,NN3,NN4,NN5);
  end;
  Ferme(Q);
end;

procedure TOF_BTMOUVSTK.ScreenKeYDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = 120 then
  begin
    TToolbarButton97(GetControl('BchercheS')).Click;
    Key := 0;
  end;
end;

Initialization
  registerclasses ( [ TOF_BTMOUVSTK ] ) ;
end.

