{***********UNITE*************************************************
Auteur  ...... : FV1 (VAUTRAIN)
Créé le ...... : 04/04/2013
Modifié le ... :   /  /
Description .. : Uses utilitaires de recalcul des PMAP
Mots clefs ... : TOF;RECALCULPMAP
*****************************************************************}
unit UtilPMAPCalcul;

interface

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
     Types,
     AglInit;

type TTypeArtio = (TtrStock,TtrVente);

Function CalculPMAP( QteEnStock: double;  PMAP: double; PA,QteReception: double): Double;
Function LectureInfoDepot(Depot : String) :  TOB;
Procedure LectureInfoDispo(Article, Depot : String ; TOBDispo : TOB);
//
function ConstituetableTemp(fNomTable : string; Article, Depot: String; Date1, Date2: TDateTime) : boolean;
procedure DropTableExist(NomTABLE: String);
procedure MiseAJourTableTempo(fNomTable : string; QteEnStock, NewPMAP,PMAPUV: Double;UniteMvt : string; QQ: TQuery);
Procedure MiseAjourDispoInv(TobMAJ : TOB;FPrix : array of String; recalcPmap : boolean);
function CalculsTableTempo(fNomTable : string; Depot, Article : String; var PmapUS,QTEStockUS : double;PmapIni:double = 0; Qteini : double = 0) : boolean;

function ConstitueNomTemp : string;
procedure ReajustePmapSuiteInv (Depot,Article : string; DateInv : TDateTime; QteInv,PmapINV : Double);
Procedure MiseAjourPMAPLiv(TobMaj,TOBTempo : TOB);
Procedure MAjPMAPConso(TobMaj : TOB);
Procedure MiseAjourDispoSuiteInv(Depot,Article : string; QteEnStock,PmapStock : double);
procedure LanceTraitementRecalcDepot (CodeDepot : string; TraitLiv,TraitConso : boolean);

implementation

Uses MSGUtil,Invutil,StockUtil,EntGC,DateUtils,
		 Utilsaisieconso,uEntCommun,FactCalc,FactTOB,
     UtilTOBPiece,PiecesRecalculs,FactOuvrage,SAISUTIL;


function ConstitueNomTemp : string;
var fDay,fMonth,fyear,fHour,fMin,fsec,fmili : word;
begin
  DecodeDateTime (Now,fyear,fMonth,fDay,fHour,fMin,fSec,fmili);
	result := 'BTW'+V_PGI.User+InttoStr(fyear)+InttoStr(fmonth)+InttoStr(fMili);
end;



function ConstituetableTemp(fNomTable : string; Article, Depot: String; Date1, Date2: TDateTime) : boolean;


  function QteSousDetail(Qte : Double ; cledoc: r_cledoc; N1, N2, N3, N4, N5: Integer): double;

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

  function QteLigne (cledoc : r_cledoc) : Double;
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

  procedure RecalculSousDetailOuv;
  var QQ : TQuery;
      Cledoc : r_cledoc;
      N1,N2,N3,N4,N5 : Integer;
      QteReelle : Double;
  begin
    // Parcours des détails d'ouvrages
    QQ := OpenSql ('SELECT * FROM '+fNomTable+' WHERE TYPEDATA="OUV"',false);
    if not QQ.eof then
    begin
      QQ.First;
      While not QQ.eof do
      begin
        //
        Cledoc.NaturePiece := QQ.findfield('NATUREPIECEG').AsString;
        Cledoc.Souche := QQ.findfield('SOUCHE').AsString ;
        Cledoc.NumeroPiece := QQ.findfield('NUMERO').AsInteger ;
        Cledoc.Indice := QQ.findfield('INDICEG').AsInteger ;
        Cledoc.NumLigne  := QQ.findfield('NOLIGNE').AsInteger ;
        //
        N1 :=QQ.findfield('N1').AsInteger;
        N2 :=QQ.findfield('N2').AsInteger;
        N3 :=QQ.findfield('N3').AsInteger;
        N4 :=QQ.findfield('N4').AsInteger;
        N5 :=QQ.findfield('N5').AsInteger;
        //
        QteReelle := QteLigne (Cledoc) * QteSousDetail( QQ.findfield('QTESTOCK').Asfloat, Cledoc,N1,N2,N3,N4,N5);
        //
        QQ.Edit;
        QQ.FindField('QTEMVT').AsFloat := QteReelle;
        QQ.UpdateRecord;
        QQ.Next;
      end;
    end;
    ferme (QQ);
  end;

  procedure SupprimeLivFromFourn;
  var QQ : TQuery;
  begin
    QQ := OpenSql ('SELECT * FROM '+fNomTable+' WHERE NATUREPIECEG IN ("LBT","BLC")',false);
    if not QQ.eof then
    begin
      QQ.First;
      While not QQ.eof do
      begin
        if not IsLivChantier (QQ.FindField('PIECEPRECEDENTE').AsString,QQ.FindField('PIECEORIGINE').AsString) then QQ.next
                                                                     else QQ.Delete;
      end;
    end;
    ferme (QQ);
  end;


var StsQl         : String;
    NB            : Integer;
    QQ : TQuery;
begin
  Nb := 0;
  DropTableExist(fNomTable);

  StSQL := 'SELECT "LIG" AS TYPEDATA,TYPELIGNE, GL_NATUREPIECEG AS NATUREPIECEG,GL_SOUCHE AS SOUCHE, ';
  StsQl := StsQl + 'GL_NUMERO AS NUMERO,GL_INDICEG AS INDICEG, GL_NUMLIGNE AS NOLIGNE, '+
                   'CAST(0 AS NUMERIC(19,4)) AS N1,'+
                   'CAST(0 AS NUMERIC(19,4)) AS N2,'+
                   'CAST(0 AS NUMERIC(19,4)) AS N3,'+
                   'CAST(0 AS NUMERIC(19,4)) AS N4,'+
                   'CAST(0 AS NUMERIC(19,4)) AS N5,'+
                   'GL_NUMORDRE AS NUMORDRE,';
  StSQL := StSQL + 'GL_DEPOT AS ONEDEPOT,GDE_LIBELLE AS LIBDEPOT, GL_ARTICLE AS ONEARTICLE,';
  StSQL := StSQL + 'GL_LIBELLE AS LIBARTICLE, GL_TIERS AS ONETIERS,T_LIBELLE AS LIBTIERS, GL_DATECREATION AS DATECREATION,GL_DATEPIECE AS DATEPIECE,';
  StSQL := StSQL + 'GL_QTESTOCK AS QTESTOCK,GL_QTEFACT AS QTEMVT,GL_DPA AS PA, PU,';
  StSQL := StSQL + 'GL_PMAP AS PMAP,CAST(0 AS NUMERIC(19,4)) AS PMAPUV,';
  StsQl := StSql + 'GPP_ESTAVOIR,UNITEACH,COEFUAUS,UNITEVTE,COEFUSUV,CAST(0 AS NUMERIC(19,4)) AS PUSTK,';
  StSql := StSql + 'CAST(0 AS NUMERIC(19,4)) AS QTESTK,(SELECT GA_QUALIFUNITESTO FROM ARTICLE WHERE GA_ARTICLE=GL_ARTICLE) AS UNITESTO,';
  StSql := StSql + 'CAST("   " AS VARCHAR(3)) AS UNITEMVT,CAST(0 AS NUMERIC(19,4)) AS MTSTK,0 AS UNIQUEBLO,';
  StSql := StSql + 'GL_PIECEORIGINE AS PIECEORIGINE,GL_PIECEPRECEDENTE AS PIECEPRECEDENTE ';
  StSQL := STSQL + 'INTO '+fNomTable+' ';
  StSQL := StSQL + 'FROM BTMOUVSTK LEFT JOIN TIERS ON T_TIERS=GL_TIERS ';
  StSQL := StSQL + 'LEFT JOIN DEPOTS ON GDE_DEPOT=GL_DEPOT WHERE ';
    //
  IF Article <> '' then StSQL := STSQL + 'GL_ARTICLE="' + Article + '" AND ';
  //
  StSQL := StSQL + 'GL_DEPOT="' + Depot +'" AND ';
  StSQL := StSQL + 'GL_QTERESTE <> 0 AND ';
  StSQl := StSQL + 'GL_DATEPIECE >="' + USDATETIME(Date1) + '" AND GL_DATEPIECE <="' + USDATETIME(Date2) + '" ';
  StSQL := StSQL + 'ORDER BY GL_ARTICLE,GL_DATEPIECE,GL_DATECREATION';

  ExecuteSql (StSql);
  StSQL := 'INSERT INTO '+fNomTable+
           '(TYPEDATA,TYPELIGNE,NATUREPIECEG,SOUCHE, NUMERO,NOLIGNE,'+
           'N1,N2,N3,N4,N5,NUMORDRE,INDICEG,ONEDEPOT,LIBDEPOT,ONEARTICLE,'+
           'LIBARTICLE,ONETIERS,LIBTIERS,DATECREATION,DATEPIECE,QTESTOCK,QTEMVT,PA,PU,PMAP,PMAPUV,GPP_ESTAVOIR,'+
           'UNITEACH,'+
           'COEFUAUS,UNITEVTE,PUSTK,QTESTK,UNITESTO,UNITEMVT,MTSTK,COEFUSUV,UNIQUEBLO,PIECEORIGINE,PIECEPRECEDENTE) '+
           'SELECT '+
           '"OUV","SORTIE",BLO_NATUREPIECEG,BLO_SOUCHE,BLO_NUMERO,BLO_NUMLIGNE,BLO_N1,BLO_N2,BLO_N3,BLO_N4,BLO_N5,'+
           'NUMORDRE,BLO_INDICEG,BLO_DEPOT,(SELECT GDE_LIBELLE FROM DEPOTS WHERE GDE_DEPOT=BLO_DEPOT),BLO_ARTICLE,BLO_LIBELLE,'+
           'BLO_TIERS,(SELECT T_LIBELLE FROM TIERS WHERE T_TIERS=BLO_TIERS AND T_NATUREAUXI="CLI"),'+
           'BLO_DATECREATION,BLO_DATEPIECE,BLO_QTEFACT,BLO_QTEFACT,BLO_DPA,BLO_PUHTDEV,BLO_PMAP,0,GPP_ESTAVOIR,'+
           'BLO_QUALIFQTEACH,BLO_COEFCONVQTE,BLO_QUALIFQTEVTE,0,0,BLO_QUALIFQTESTO,"   ",0,BLO_COEFCONVQTEVTE,BLO_UNIQUEBLO,'+
           'GL_PIECEORIGINE AS PIECEORIGINE,GL_PIECEPRECEDENTE AS PIECEPRECEDENTE '+
           'FROM BTLISTMSTOCKO WHERE ';
  if Article <> '' Then StsQl := StsQl + 'BLO_ARTICLE="'+Article+'" AND ';
  StSQL := StSQL + 'BLO_DEPOT="' + Depot +'" ';
  StSQl := StSQL + 'AND BLO_DATEPIECE >="' + USDATETIME(Date1) + '" AND BLO_DATEPIECE <="' + USDATETIME(Date2) + '" ';
  StSQL := StSQL + 'ORDER BY BLO_ARTICLE,BLO_DATEPIECE,BLO_DATECREATION';
  ExecuteSql (StSql);
  ExecuteSql ('CREATE INDEX OOOOINDEX1 ON '+fNomTable+' (ONEDEPOT,ONEARTICLE,DATEPIECE,DATECREATION)');
  ExecuteSql ('CREATE INDEX OOOOINDEX2 ON '+fNomTable+' (TYPEDATA)');
  SupprimeLivFromFourn;
  RecalculSousDetailOuv;

  QQ := OpenSQL('SELECT COUNT(*) FROM '+fNomTable,True,1,'',true );
  if not QQ.eof then Nb := QQ.fields[0].AsInteger;
  ferme (QQ);
  Result := (Nb > 0);
end;


Procedure MAjPMAPConso(TobMaj : TOB);
Var DPA     : Double;
    DPR     : Double;
    PUHT    : Double;
    CoefPR  : Double;
    CoefPU  : Double;
    //
    Q3      : TQuery;
    stSql : string;
    TOBConso : TOB;
begin
  TOBConso := TOB.Create ('CONSOMMATIONS',nil,-1);
  StSQL  := 'SELECT * FROM CONSOMMATIONS WHERE BCO_NATUREPIECEG="' + TOBMAJ.GetString('NATUREPIECEG') +
            '"  AND BCO_SOUCHE="'+ TOBMAJ.GetString('SOUCHE') +
            '"  AND BCO_NUMERO=' + TOBMAJ.GetString('NUMERO') +
            '   AND BCO_NUMORDRE=' + TOBMAJ.GetString('NUMORDRE');

  Q3 := OpenSQL(StSQL, True);
  TRY
    If Not Q3.eof then
    begin
      TobConso.SelectDB('',Q3,False);
      DPA := TobConso.GetDouble('BCO_DPA');
      DPR := TobConso.GetDouble('BCO_DPR');
      PUHT:= TobConso.GetDouble('BCO_PUHTDEV');
      TobConso.PutValue('BCO_DPA', TobMaj.GetDouble('PMAP'));
      //Recalcul du DPR et du PUHTDEV en fonction du PMAP Si Pièce de type LBT
      if DPA <> 0 then CoefPR  := (DPR/DPA) else Coefpr := 0;
      if DPR <> 0 then CoefPU  := (PUHT/DPR) else CoefPU := 0;
      // Calcul de la ligne
      Utilsaisieconso.calculeLaLigne(TOBConso,CoefPR,CoefPU);
      //
      Tobconso.UpdateDB;
    end;
  FINALLY
    Ferme(Q3);
    TOBConso.free;
  END;
end;


Procedure MiseAjourPMAPLiv(TobMaj,TOBTempo : TOB);
Var DPA     : Double;
    DPR     : Double;
    PMRP    : Double;
    PMAP    : Double;
    Q2      : TQuery;
    QQ      : TQuery;
    cledoc  : R_Cledoc;
    stSql : string;
    TOBLiv : TOB;
    TobLTempo : TOB;
    DEV : Rdevise;
begin
  if TobMaj.getInteger('UNIQUEBLO') = 0 then
  BEGIN
    TOBLIv := TOB.Create ('LIGNE',nil,-1);
  StSQL  := 'SELECT * FROM LIGNE WHERE '  +
            '       GL_NATUREPIECEG="'    + TOBMAJ.GetString('NATUREPIECEG') +
            '"  AND GL_SOUCHE="'          + TOBMAJ.GetString('SOUCHE') +
            '"  AND GL_NUMERO='           + TOBMAJ.GetString('NUMERO') +
            '   AND GL_NUMLIGNE='         + TOBMAJ.GetString('NOLIGNE');

  Q2 := OpenSQL(StSQL, True);

  TRY
    If Not Q2.eof then
    begin
      //
      TobLiv.SelectDB('',Q2,False);
      DPA := TobLiv.GetDouble('GL_DPA');
      DPR := TobLiv.GetDouble('GL_DPR');
  //      PUHT:= TobLiv.GetDouble('GL_PUHTDEV');
      PMAP:= TobMaj.GetDouble('GL_PMAP');
      PMRP:= TobMaj.GetDouble('GL_PMRP');
      //
      TobLiv.PutValue('GL_DPA', TobMaj.GetDouble('PMAPUV'));
      TobLiv.PutValue('GL_COEFMARG', 0);
      //
      CalculeMontantLigneAchat(TOBLiv, DPA, DPR,PMAP, PMRP);
      //
      TobLiv.UpdateDB;

      //charge tob Tempo avec mémo clepiece
      if TOBTempo <> nil then
      begin
        TobLTempo := TOBTempo.FindFirst(['GP_NATUREPIECEG','GP_SOUCHE','GP_NUMERO','GP_INDICEG'], [TOBLiv.GetString('GL_NATUREPIECEG'), TOBLiv.GetString('GL_SOUCHE'), TOBLiv.GetString('GL_NUMERO'), TOBLiv.GetString('GL_INDICEG')], false);
        if TOBLTempo = nil then
        begin //si non création enregistrement
          //lecture de la pièce
          cledoc := TOB2cledoc (TOBLIV);
          StSQl   := 'SELECT * FROM PIECE WHERE '+ WherePiece (Cledoc,ttdpiece,false);
          QQ := OpenSql(StSQL,false);
          if not QQ.eof then
          begin
            TOBLTempo := TOB.Create('PIECE', TOBTempo, -1);
            TOBLTempo.selectDb('', QQ);
          end;
    			Ferme(QQ);
        end;
      end;
    end;
  FINALLY
    Ferme(Q2);
    TOBLiv.free;
  END;
  END ELSE
  BEGIN
    TOBLIv := TOB.Create ('LIGNEOUV',nil,-1);
    StSQL  := 'SELECT * FROM LIGNEOUV WHERE '  +
              'BLO_NATUREPIECEG="'    + TOBMAJ.GetString('NATUREPIECEG') + '" ' +
              'AND BLO_SOUCHE="'+ TOBMAJ.GetString('SOUCHE') +'" '+
              'AND BLO_NUMERO='+ TOBMAJ.GetString('NUMERO') +' '+
              'AND BLO_NUMLIGNE='+ TOBMAJ.GetString('NOLIGNE')+' '+
              'AND BLO_UNIQUEBLO='+ TOBMAJ.GetString('UNIQUEBLO');

    Q2 := OpenSQL(StSQL, True);

    TRY
      If Not Q2.eof then
      begin
       //
        TobLiv.SelectDB('',Q2,False);
        DEV.Code := TOBLiv.GetString('BLO_DEVISE');
        GetInfosDevise(DEV);
        DEV.Taux := GetTaux(DEV.Code, DEV.DateTaux, TOBLiv.GetValue('BLO_DATEPIECE'));
        DPA := TobLiv.GetDouble('BLO_DPA');
        DPR := TobLiv.GetDouble('BLO_DPR');
  //      PUHT:= TobLiv.GetDouble('GL_PUHTDEV');
        PMAP:= TobMaj.GetDouble('BLO_PMAP');
        PMRP:= TobMaj.GetDouble('BLO_PMRP');
        //
        TobLiv.PutValue('BLO_DPA', TobMaj.GetDouble('PMAPUV'));
        TobLiv.PutValue('BLO_COEFMARG', 0);
        //
        CalculeLigneAcOuv(TOBLiv,DEV,false);
        //
        TobLiv.UpdateDB;

        //charge tob Tempo avec mémo clepiece
        if TOBTempo <> nil then
        begin
          TobLTempo := TOBTempo.FindFirst(['GP_NATUREPIECEG','GP_SOUCHE','GP_NUMERO','GP_INDICEG'], [TOBLiv.GetString('BLO_NATUREPIECEG'), TOBLiv.GetString('BLO_SOUCHE'), TOBLiv.GetString('BLO_NUMERO'), TOBLiv.GetString('BLO_INDICEG')], false);
          if TOBLTempo = nil then
          begin //si non création enregistrement
            //lecture de la pièce
            cledoc := TOB2cledoc (TOBLIV);
            StSQl   := 'SELECT * FROM PIECE WHERE '+ WherePiece (Cledoc,ttdpiece,false);
            QQ := OpenSql(StSQL,false);
            if not QQ.eof then
            begin
              TOBLTempo := TOB.Create('PIECE', TOBTempo, -1);
              TOBLTempo.selectDb('', QQ);
            end;
            Ferme(QQ);
          end;
        end;
      end;
    FINALLY
      Ferme(Q2);
      TOBLiv.free;
    END;
  END;

end;

function CalculsTableTempo(fNomTable : string; Depot, Article : String; var PmapUS,QTEStockUS : double;PmapIni:double = 0; Qteini : double = 0) : boolean;

	function GetRatioDoc (QQ : TQuery; TypeRatio : TTypeArtio ) : double;
  var VenteAchat : string;
  		CoefUaUs,CoefUSUv : double;
      UniteAch,UniteVte,UniteSto : string;
  begin
    result := 1;
		VenteAchat:=GetInfoParPiece(QQ.findField('NATUREPIECEG').AsString,'GPP_VENTEACHAT') ;
    CoefUAUS := QQ.findfield('COEFUAUS').AsFloat;
    CoefUSUV := QQ.findfield('COEFUSUV').AsFloat;
    UniteAch := trim(QQ.findField('UNITEACH').AsString);
    UniteSto := Trim(QQ.findField('UNITESTO').AsString);
    UniteVte := Trim(QQ.findField('UNITEVTE').AsString);
    //
    Case TypeRatio of

      TtrStock :
      BEGIN
        if VenteAchat='ACH' then
        BEGIN
          if CoefuaUs <> 0 then
          begin
            result := 1/CoefUaUs;
          end else
          begin
            result:=CalcRatioVA(UniteAch,UniteSto) ;
          end;
        END else if (VenteAchat='AUT') or (VenteAchat='TRF') then
        BEGIN
          result:=1.0 ;
        END else
        BEGIN
          if CoefUSUv <>0 then
          begin
            result:= CoefusUv;
          end else
          begin
            result:=CalcRatioVA(UniteVte,UniteSto) ;
          end;
        END ;
      END ;

      TtrVente :
      BEGIN
        if VenteAchat='ACH' then
        BEGIN
          if CoefUaUs <> 0 then
          begin
            if CoefUSUV <> 0 then
            begin
              result := CoefUSUv;
            end else
            begin
              result:=CalcRatioVA(UniteVte,UniteSto) ;
            end;
            result := result * CoefUaUs;
          end else
          begin
            result:=CalcRatioVA(UniteVte,UniteAch) ;
          end;
        END else if (VenteAchat='AUT') or (VenteAchat='TRF') then
        BEGIN
          result:=1.0 ;
        END else
        BEGIN
          result:=CalcRatioVA(UniteVte,UniteSto) ;
        END ;
      END ;
    END ;
  end;

	function PrxAch2US (QQ : TQuery; Prix : Double) : Double;
  var RatioVA : Double;
  begin
    RatioVa := GetRatioDoc(QQ,TtrStock);
    Result := ARRONDI(Prix*RatioVa,V_PGI.okdecP);
  end;

  function QteAch2US (QQ : TQuery; Qte : Double) : Double;
  var RatioVA : Double;
  begin
    RatioVa := GetRatioDoc(QQ,TtrStock);
    Result := ARRONDI(Qte/RatioVa,V_PGI.okdecP);
  end;

  function QteVte2US (QQ : TQuery; Qte : double) : Double;
  var RatioVA : Double;
  begin
    RatioVa := GetRatioDoc(QQ,TtrStock);
    Result := ARRONDI(Qte/RatioVa,V_PGI.okdecP);
  end;

  function MtSTo2UV (QQ : TQuery ; Prix : Double) : Double;
  var RatioVA : Double;
  begin
    RatioVa := GetRatioDoc(QQ,TtrVente);
    Result := ARRONDI(Prix*RatioVa,V_PGI.okdecP);
  end;

var StSQl,UniteMvt         : String;
    //
    QQ            : Tquery;
    //
    NewPmap,PMAPUV : Double;
    QteEnStock    : Double;
    QteSortie     : Double;
    QteReception  : Double;
    PA            : Double;
    PAUS,QTEUS : double;
    //
    TobDispo      : TOB;
    zDep,zArt : string;
begin
  //
  result := false;
  zdep := '';
  zArt := '';
  QteEnStock := 0;
  NewPmap := 0;
  PMAPUV := 0;
  //
  TobDispo := TOB.Create('DISPO',nil,-1);
  //'PMAP et QteStock du stock lors de l''inventaire' (GQ_PMAPINV, GQ_STOCKINV)
  //lecture de la table tempo pour recalcul PMAP
  StSql := 'SELECT * FROM '+fNomTable+' ORDER BY ONEDEPOT,ONEARTICLE,DATEPIECE,DATECREATION';
  QQ := OpenSQL(StSQL, False);
  if not QQ.eof then
  begin
    result := True;
    QQ.First;

    while Not QQ.EOF do
    begin
      if (zdep <> QQ.FindField('ONEDEPOT').Asstring) or (zart <> QQ.FindField('ONEARTICLE').Asstring)  then
      begin
        zDep     := QQ.FindField('ONEDEPOT').Asstring;
        zArt    := QQ.FindField('ONEARTICLE').Asstring;
        if (PmapIni <> 0) or (Qteini <> 0) then
        begin
          NewPmap    := ARRONDI(PmapIni,V_PGI.OkDecP); // --> Forcement en US
          QteEnStock := ARRONDI(Qteini,V_PGI.OkdecQ); // --> forcement en US
        end else
        begin
        LectureInfoDispo(Article, Depot, TOBDispo);
        NewPmap    := ARRONDI(TobDispo.GetDouble('GQ_PRIXINV'),V_PGI.OkDecP); // --> Forcement en US
        QteEnStock := ARRONDI(Tobdispo.GetDouble('GQ_STOCKINV'),V_PGI.OkdecQ); // --> forcement en US
      end;
      end;

      IF QQ.FindField('TYPELIGNE').Asstring = 'ENTREE' then
      begin
        QteReception  := QQ.FindField('QTEMVT').AsFloat ;
        PA            := QQ.FindField('PU').AsFloat ;
        // --> passage des infos en US
        PAUS := PrxAch2Us(QQ,PA);
        QTEUS := QteAch2US(QQ,QteReception);
        //
        if PA <> 0 then NewPMAP := CalculPMAP(QteEnStock,NewPmap,PAUS,QTEUS);
        //
        QteEnStock    := QteEnStock+QTEUS;
        UniteMvt := QQ.FindField('UNITEACH').AsString;
      end else if QQ.FindField('TYPELIGNE').AsString = 'SORTIE' then
      begin
        //Recalcul du stock
        PMAPUV := MtSTo2UV (QQ,NewPmap);
        QteSortie     := QQ.FindField('QTEMVT').AsFloat ;
        // --> passage des infos en US
        QTEUS := QteVte2US(QQ,QteSortie);
        //
        QteEnStock    := QteEnStock-QTEUS;
        UniteMvt := QQ.FindField('UNITEVTE').AsString;
      end;

      MiseAJourTableTempo(fNomTable,QteEnStock, NewPMAP, PMAPUV,UniteMvt,QQ);

      QQ.Next;
    end;
  end;
  ferme (QQ);
  TobDispo.Free;
  //
	PmapUS := NewPmap;
	QTEStockUS := QteEnStock;
end;

procedure MiseAJourTableTempo(fNomTable : string; QteEnStock, NewPMAP,PmapUV: Double;UniteMvt : string; QQ: TQuery);
var StSQL : String;
    Nb    : Integer;
    MTSTK : double;
begin
	MTSTK := ARRONDI(QteEnStock * NewPmap,V_PGI.OkDecV); if MTSTK < 0 then MTSTK := 0;
  StsQL :='UPDATE '+fNomTable+' ';
  StSQL := StSQL + 'SET QTESTOCK=' + StrfPoint(Arrondi(QteEnStock,V_PGI.okdecQ)) + ',';
  StSQL := StSQL + 'PMAP=' + StrfPoint(arrondi(NewPMAP,V_PGI.OkdecP))+',';
  StSQL := StSQL + 'PMAPUV=' + StrfPoint(Arrondi(PMAPUV,V_PGI.okdecP))+',';
  StSQL := StSQL + 'MTSTK=' + StrfPoint(Arrondi(MTSTK,V_PGI.okdecP))+',';
  StSQL := StSQL + 'UNITEMVT="'+ trim(UniteMvt)+'" ';
  StSQL := StSQL + 'FROM '+fNomTable+' ';
  StSQL := StSQL + 'WHERE NATUREPIECEG="'  + QQ.FindField('NATUREPIECEG').AsString + '"';
  StSQL := StSQL + '   AND SOUCHE="'        + QQ.FindField('SOUCHE').AsString + '"';
  StSQL := StSQL + '   AND NUMERO="'        + QQ.FindField('NUMERO').AsString + '"';
  StSQL := StSQL + '   AND NOLIGNE="'       + QQ.FindField('NOLIGNE').AsString + '"';
  StSQL := StSQL + '   AND ONEDEPOT="'         + QQ.FindField('ONEDEPOT').AsString + '"';
  StSQL := StSQL + '   AND ONEARTICLE="'       + QQ.FindField('ONEARTICLE').AsString + '"';

  Nb := ExecuteSQL(StSQL);
end;

function CalculPMAP(QteEnStock : double; PMAP : double;PA, QteReception : double) : Double;
begin
  Result := 0;
  if QteEnStock < 0 then BEGIN Result := PA; Exit; end;
  //Calcul du PMAP = (Pu*QtéStock)+(PA*QtéRéception)/(QtéStock+QtéRéception)
  if (QteEnStock+QteReception) <> 0 then
    Result := Arrondi(((PMAP*QteEnStock)+(PA*QteReception))/(QteEnStock+QteReception),V_PGI.okdecP);
end;


Procedure DropTableExist(NomTABLE : String);
Var StSQL : String;
    QQ    : TQuery;
begin

  StSQL := 'SELECT OBJECT_ID(N"'+ NomTable+'") AS Object_ID';
  QQ := OpenSQL(StSQL, False);

  if QQ.FindField('OBJECT_ID').AsString <> '' Then ExecuteSQL('DROP TABLE ' + NomTable);

  ferme (QQ);
end;

//lecture de la table Dépot pour affichage des éléments nécessaires
Function LectureInfoDepot(Depot : String) :  TOB;
var StSQl : String;
    QQ    : TQuery;
begin

  Result:= Tob.create('DEPOTS', nil, -1);

  StSQL := 'Select * FROM DEPOTS WHERE GDE_DEPOT="' + Depot + '"';
  QQ := OpenSQL(StSQL, False);
  If Not QQ.Eof then
  begin
    Result.SelectDB('',QQ,False);
  end;

  ferme (QQ);

end;

//lecture de la table DISPO (GD_) pour affichage des éléments de Stocks
procedure LectureInfoDispo(Article, Depot : String; TOBdispo : TOB);
var StSQl : String;
    QQ    : TQuery;
begin
  TOBdispo.InitValeurs(false);
  StSQL := 'Select * FROM DISPO WHERE GQ_ARTICLE="' + Article + '" AND GQ_DEPOT="' + Depot + '"';
  QQ := OpenSQL(StSQL, False);
  If Not QQ.Eof then
  begin
    TOBdispo.SelectDB('',QQ,False);
  end;

  ferme (QQ);

end;

function ChoozePrice(BaseChamp : String; TOBInvLig : TOB; Fprix : array of string) : Double;
var i : integer;
begin
  result := 0;
  for i := 0 to 2 do
  begin
     result := TOBInvLig.GetDouble(BaseChamp + RechDom('GCTYPEPRIX', FPrix[I], true));
     if Result <> 0 then Break;
  end;
end;

procedure EnregistreHistoInv(TOBDispo : TOB; DateInv : TDateTime);
var TOBH : TOB;
    StSQl : string;
    OneDate,CmpDate : TdateTime;
    QQ : TQuery;
begin
  CmpDate :=  StrToDate(DateToStr(DateInv));
  OneDate := StrToDate(DateToStr(TOBDispo.GetDateTime ('GQ_DATEINV')));
  if OneDate = CmpDate then exit;
  TOBH := TOB.Create ('BHISTOINV',nil,-1);
  TRY
    StSQL  := 'SELECT * FROM BHISTOINV WHERE '+
              'BHI_DEPOT="' + TOBDispo.GetString('GQ_DEPOT') +'" AND '+
              'BHI_ARTICLE="'+ TOBDispo.GetString('GQ_ARTICLE') + '" AND ' +
              'BHI_DATEINV="'+ USDateTime(OneDate) + '"';
    QQ := OpenSql (StSql,true,1,'',true);
    if not QQ.eof then
    begin
      TOBH.SelectDB ('',QQ);
      TOBH.AddChampSupValeur('NEW','-');
    end else
    begin
      TOBH.AddChampSupValeur('NEW','X');
      TOBH.SetString('BHI_DEPOT',TOBDispo.GetString('GQ_DEPOT'));
      TOBH.SetString('BHI_ARTICLE',TOBDispo.GetString('GQ_ARTICLE'));
      TOBH.SetDateTime('BHI_DATEINV',OneDate);
      TOBH.SetAllModifie(true);
    end;
    ferme (QQ);
    if TOBDispo.GetDouble('GQ_STOCKINV')<>0 then TOBH.SetDouble('BHI_STOCKINV',TOBDispo.GetDouble('GQ_STOCKINV'))
                                            else TOBH.SetDouble('BHI_STOCKINV',TOBDispo.GetDouble('GQ_PHYSIQUE'));
    if TOBDispo.GetDouble('GQ_PRIXINV') <> 0 then TOBH.SetDouble('BHI_PRIXINV',TOBDispo.GetDouble('GQ_PRIXINV'))
                                             else TOBH.SetDouble('BHI_PRIXINV',TOBDispo.GetDouble('GQ_PMAP'));
    if TOBH.getString('NEW')='X' then
    begin
      TOBH.InsertDB(nil);
    end else
    begin
      TOBH.UpdateDB(false); 
    end;
  FINALLY
    TOBH.free;
  END;
end;

Procedure MiseAjourDispoInv(TobMAJ : TOB;FPrix : array of String; recalcPmap : boolean);
Var DateInv   : TDateTime;
    PMAPInv   : Double;
    QteInv    : Double;
    DPA,DPR       : Double;
    //
    StSQL,Article     : String;
    //
    Q4        : TQuery;
    TOBDispo : TOB;
    KKSuite : Boolean;
begin
  KKsuite := False;
  Article := TOBMAJ.GetString('GIL_ARTICLE');
  TOBDispo := TOB.Create ('DISPO',nil,-1);
  //
  PMAPInv   := ChoozePrice('GIL_PMAP', TobMaj,FPrix );
  DPA       := ChoozePrice('GIL_DPA', TobMaj,FPrix );
  DPR       := ChoozePrice('GIL_DPR', TobMaj,FPrix );
  QteInv    := TobMaj.Getdouble('GIL_INVENTAIRE');
  DateInv   := TOBMaj.GetDateTime('GIL_DATESAISIE');
  //
  StSQL  := 'SELECT * FROM DISPO WHERE GQ_DEPOT="' + TOBMAJ.GetString('GIL_DEPOT') +
            '"  AND GQ_ARTICLE="'+ TOBMAJ.GetString('GIL_ARTICLE') + '"';

  Q4 := OpenSQL(StSQL, false);

  If Not Q4.eof then
  begin
    TOBDispo.SelectDB('',Q4);
    EnregistreHistoInv(TOBDispo,DateInv);
    //
    TOBDispo.SetDouble ('GQ_PRIXINV',PMAPInv);
    TOBDispo.setDouble ('GQ_PMAP', PMAPINV);
    //FV1 : 15/02/2016 - FS#1857 - VEUVE CHATAIN - problème de cohérence des quantités en stocks
    TOBDispo.setDouble ('GQ_STOCKINV', QteInv);
    TOBDispo.setDouble ('GQ_PHYSIQUE', QteInv);
    //FV1 : 15/02/2016 - FS#1857 - VEUVE CHATAIN - problème de cohérence des quantités en stocks
    TOBDispo.SetDateTime ('GQ_DATEINV',DateInv);
    TOBDispo.setDouble ('GQ_DPA', DPA);
    TOBDispo.setDouble ('GQ_DPR', DPR);
    TOBDispo.setAllModifie(true);
    TOBDispo.UpdateDB(false);
    KKSuite := true;
  end;

  Ferme(Q4);
  TOBDispo.free;

  if (KKSuite) and (recalcPmap) then
  begin
    ReajustePmapSuiteInv (TobMaj.GetString('GIL_DEPOT'),TobMaj.GetString('GIL_ARTICLE'),DateInv,QteInv,PMAPInv);
  end;
end;

Procedure MiseAjourDispoSuiteInv(Depot,Article : string; QteEnStock,PmapStock : double);
Var StSql : string;
		QQ : TQuery;
    TT : TOB;
begin
	TT := TOB.Create('DISPO',nil,-1);
  StSQL  := 'SELECT * FROM DISPO WHERE GQ_DEPOT="' + Depot + '"  AND GQ_ARTICLE="'+ Article + '"';
  QQ := OpenSQL(StSql,True,1,'',true);
  if not QQ.eof then
  begin
    TT.selectDb ('',QQ);
    TT.SetDouble('GQ_PMAP',Arrondi(PmapStock,V_PGI.okdecQ));
    TT.SetDouble('GQ_PHYSIQUE',Arrondi(QteEnStock,V_PGI.okdecQ));
    TT.UpdateDB(false);
  end;
  Ferme(QQ);
  TT.free;
end;


procedure ReajustePmapSuiteInv (Depot,Article : string; DateInv : TDateTime; QteInv,PmapINV : Double);
var fNomTable,StSql : string;
		II : Integer;
    QQ : TQuery;
    TOBDet,TOBPRecalc : TOB;
    PMAPStock,QteStock : double;
begin
  TOBPRecalc := TOB.Create('LES PIECES',nil,-1);
  TOBDet := TOB.Create('UN DETAIL',nil,-1);
	fNomTable := ConstitueNomTemp;
	ConstituetableTemp (fNomTable,Article,Depot,DateInv,Now);
  TRY
    if CalculsTableTempo(fNomTable,Depot, Article,PMAPStock,QteStock) then
    begin
      //lecture de la table tempo pour reajuster le PMAP
      StSql := 'SELECT * FROM '+fNomTable+' ORDER BY ONEDEPOT,ONEARTICLE,DATEPIECE,DATECREATION';
      QQ := OpenSQL(StSQL, False);
      QQ.First;
      while Not QQ.EOF do
      begin
        IF (QQ.findField('NATUREPIECEG').AsString='LBT') OR (QQ.findField('NATUREPIECEG').AsString='BLC') then
        begin
          TOBDet.SelectDB('',QQ);
          MiseAjourPMAPLiv (TOBDet,TOBPrecalc);
          MAjPMAPConso (TOBDet);
        end;
        QQ.Next;
      end;
      ferme (QQ);
      if TOBPrecalc.detail.Count > 0 then
      begin
        For II := 0 to TOBPrecalc.detail.count-1 do
        begin
          if TraitementRecalculPiece(TOBPrecalc.Detail[II],false,false)<>TrrOk  then V_PGI.ioerror := OeUnknown;
        end;
      end;
    end else
    begin
    	QteStock := QteInv;
      PMAPStock := PmapINV;
    end;
    MiseAjourDispoSuiteInv (Depot,Article,QteStock,PmapStock);
    //
  FINALLY
	 	DropTableExist(fNomtable);
    TOBDet.Free;
    TOBPRecalc.free;
  end;
end;

procedure LanceTraitementRecalcDepot (CodeDepot : string; TraitLiv,TraitConso : boolean);
var QQ : TQuery;
		StSql : string;
    TOBDispos,TOBDispo,TOBDet,TOBPieces : TOB;
    II : Integer;
    fNomTable,Article,Depot : string;
    DateInv : TDateTime;
    PmapStock,QteStock,PmapInv,QteInv : double;
begin
  fNomTable := ConstitueNomTemp;
  //
  TOBDispos := TOB.Create('LES DISPOS',nil,-1);
  TOBDet := TOB.Create ('LES TEMPO',nil,-1);
  TOBPieces := TOB.Create ('LES TEMPO',nil,-1);
  //
  StSql := 'SELECT GQ_ARTICLE,GQ_DEPOT,GQ_DATEINV,GQ_STOCKINV,GQ_PRIXINV FROM DISPO '+
  				 'LEFT JOIN ARTICLE ON GA_ARTICLE=GQ_ARTICLE '+
           'WHERE GA_TYPEARTICLE IN ("MAR","ARP") AND GQ_DEPOT="'+CodeDepot+'"';
  //
  QQ := OpenSQL(StSql,True,-1,'',true);
  TOBDispos.LoadDetailDB('DISPO','','',QQ,false);
  ferme (QQ);
  TRY
    for II := 0 to TOBDispos.detail.count -1 do
    begin
      TOBDispo := TOBDispos.detail[II];
      Article := TOBDispo.GetString('GQ_ARTICLE');
      Depot := TOBDispo.GetString('GQ_DEPOT');
      DateInv := TOBDispo.GetDateTime('GQ_DATEINV');
      PmapInv := TOBDispo.Getdouble('GQ_PRIXINV');
      QteInv := TOBDispo.Getdouble('GQ_STOCKINV');
      DropTableExist (fNomtable);
      ConstitueTableTemp(fNomTable,Article,Depot,DateInv,Now);
      //
      if CalculsTableTempo(fNomTable,Depot, Article,PmapStock,QteStock) then
      begin
        //lecture de la table tempo pour reajuster le PMAP
        StSql := 'SELECT * FROM '+fNomTable+' ORDER BY ONEDEPOT,ONEARTICLE,DATEPIECE,DATECREATION';
        QQ := OpenSQL(StSQL, False);
        QQ.First;
        while Not QQ.EOF do
        begin
          IF (QQ.findField('NATUREPIECEG').AsString='LBT') OR (QQ.findField('NATUREPIECEG').AsString='BLC') then
          begin
            TOBDet.SelectDB('',QQ);
            if TraitLiv then MiseAjourPMAPLiv (TOBDet,TOBPieces);
            if TraitConso then MAjPMAPConso (TOBDet);
          end;
          QQ.Next;
        end;
        ferme (QQ);
      end else
      begin
        QteStock := Qteinv;
        PmapStock := PMAPinv;
      end;
      MiseAjourDispoSuiteInv (Depot,Article,QteStock,PmapStock);
      DropTableExist(fNomtable);
    end;
    //
    if TOBPieces.detail.Count > 0 then
    begin
      For II := 0 to TOBPieces.detail.count-1 do
      begin
        if TraitementRecalculPiece(TOBPieces.Detail[II],false,false)<>TrrOk  then V_PGI.ioerror := OeUnknown;
      end;
    end;
  FINALLY
    TOBDispos.free;
    TOBDet.free;
    TOBPieces.Free;
  END;
end;

end.
