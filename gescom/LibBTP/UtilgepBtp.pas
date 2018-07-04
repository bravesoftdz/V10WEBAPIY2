unit UtilgepBtp;

interface
uses Classes,UTOB,FactComm,Hctrls,HEnt1,SaisUtil,UtilPGi,forms,SysUtils,variants,Dialogs,Inifiles,
{$IFDEF EAGLCLIENT}
  maineagl,UtileAGL,
{$ELSE}
  Doc_Parser,DBCtrls, Db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} fe_main,uPDFBatch,EdtREtat,MajTable
{$ENDIF}
  ,M3FP
  ,Menus
	,HmsgBox
  ,HTB97
  ,Hpanel
  ,Graphics
  ,Grids
  ,Vierge
  ,Windows
  ,AglInit
  ,Messages
  ,ParamSoc
  ,Controls
  ,EntGc,uEntCommun
  ,uRecupSQLModele
  ,UtilxlsBTP
  ,FactAdresse
  ,Mul
  ,CBPPath
  ,BTGESTERROREXP_TOF
;

implementation

uses FactOuvrage,NomenUtil,FactTOB,UtilTOBPiece,UGestionErreurExportLSE,TiersUtil;
const
  TabLse = #$09;

// fonctions GEP Interne

Procedure DefinieStructureOuv (TOBGroupeNomen,TOBLN : TOB ; LaLig : integer) ;
Var TOBPere : TOB ;
    LigneN1,LigneN2,LigneN3,LigneN4,LigneN5: integer;
BEGIN
  LigneN1 := TOBLN.GetValue('BLO_N1');
  LigneN2 := TOBLN.GetValue('BLO_N2');
  LigneN3 := TOBLN.GetValue('BLO_N3');
  LigneN4 := TOBLN.GetValue('BLO_N4');
  LigneN5 := TOBLN.GetValue('BLO_N5');
  //
  if LigneN5 > 0 then
  begin
    // recherche du pere au niveau 4
    TOBPere:=TOBGroupeNomen.FindFirst(['BLO_NUMLIGNE','BLO_N1','BLO_N2','BLO_N3','BLO_N4','BLO_N5'],[LaLig,LigneN1,LigneN2,LigneN3,LigneN4,0],True) ;
  end else if TOBLN.GetValue('BLO_N4') > 0 then
  begin
    // recherche du pere au niveau 3
    TOBPere:=TOBGroupeNomen.FindFirst(['BLO_NUMLIGNE','BLO_N1','BLO_N2','BLO_N3','BLO_N4','BLO_N5'],[LaLig,LigneN1,LigneN2,LigneN3,0,0],True) ;
  end else if TOBLN.GetValue('BLO_N3') > 0 then
  begin
    // recherche du pere au niveau 2
    TOBPere:=TOBGroupeNomen.FindFirst(['BLO_NUMLIGNE','BLO_N1','BLO_N2','BLO_N3','BLO_N4','BLO_N5'],[LaLig,LigneN1,LigneN2,0,0,0],True) ;
  end else if TOBLN.GetValue('BLO_N2') > 0 then
  begin
     // recherche du pere au niveau 1
    TOBPere:=TOBGroupeNomen.FindFirst(['BLO_NUMLIGNE','BLO_N1','BLO_N2','BLO_N3','BLO_N4','BLO_N5'],[LaLig,LigneN1,0,0,0,0],True) ;
  end else TOBPere:=TOBGroupeNomen;

  if TOBPere<>Nil then
  BEGIN
    TOBLN.changeParent(TOBPere,-1) ;
  END ;
END ;

procedure BTPChargeOuvrages (cledoc : r_cledoc;TOBOuvrages : TOB; TOBPiece : TOB=nil) ;
var TOBNomen,TOBGroupenomen,TOBLN,TOBL : TOB;
    i : integer;
    Q : TQuery;
    old,Lig : integer;
    req : string;
    DEV:Rdevise;
begin
  TOBGroupeNomen := nil;
  TOBNomen:=TOB.Create('',Nil,-1) ;
  Req := 'SELECT *,BNP_TYPERESSOURCE FROM LIGNEOUV '+
         'LEFT JOIN NATUREPREST ON BNP_NATUREPRES=(SELECT GA_NATUREPRES FROM ARTICLE WHERE GA_ARTICLE=BLO_ARTICLE) '+
         'WHERE '+WherePiece(CleDoc,ttdOuvrage,False)+' ORDER BY BLO_NUMLIGNE,BLO_N1, BLO_N2, BLO_N3, BLO_N4,BLO_N5';
  Q:=OpenSQL(req,True) ;
  TOBNomen.LoadDetailDB ('LIGNEOUV','','',Q,True,true);
  Ferme(Q) ;
  if TOBNomen.detail.count <> 0 then
  begin
    i:=0;
    Old := 0;
    repeat
      TOBLN:=TOBNomen.Detail[i] ;
      InsertionChampSupOuv (TOBLN,false);
      TOBLN.putvalue('ANCPV',TOBLN.Getvalue('BLO_PUHTDEV'));
      TOBLN.putvalue('ANCPA',TOBLN.Getvalue('BLO_DPA'));
      TOBLN.putvalue('ANCPR',TOBLN.Getvalue('BLO_DPR'));
      Lig:=TOBLN.GetValue('BLO_NUMLIGNE') ;
      if Lig<>Old then
      begin
        TOBGroupeNomen := TOBOuvrages.findFirst(['BLO_NUMLIGNE'],[Lig],true);
        Old := lig;
      end;
      if TOBGroupeNomen= nil then
      BEGIN
        TOBGroupeNomen:=TOB.Create('',TOBOuvrages,-1) ;
        TOBGroupeNomen.AddChampSup('UTILISE',False) ;
        TOBGroupeNomen.PutValue('UTILISE','X') ;
        if TOBPiece <> nil then
        begin
          TOBL := TOBPiece.findFirst(['GL_NUMLIGNE'],[Lig],True);
          if TOBL <> nil then TOBL.putValue('GL_INDICENOMEN',TOBOUvrages.detail.count);
        end;
      END ;
      DefinieStructureOuv(TOBGroupeNomen,TOBLN,Lig);
      CalculMontantHtDevLigOuv (TOBLN,DEV);
    until (TOBNomen.detail.count = 0 ) or (i>TOBNomen.detail.count);
  end;
  TOBNomen.Free ;
end;

procedure BTPLoadLapiece (Cledoc : r_cledoc; TOBpiece : TOB);
begin
  TOBpiece.clearDetail;
  TOBPiece.InitValeurs;
  LoadPieceLignes (Cledoc,TOBpiece,false);
end;

procedure MiseAPlatOuvrageLocal (TOBL: TOB;Qte,QteDuDetail : double; TOBSource,TOBPlat : TOB; DEV : Rdevise);
var Indice : integer;
    QteSuite,QTeDDSuite : double;
    TOBLoc,TOBINS : TOB;
begin
  for Indice := 0 to TOBSource.detail.count -1 do
  begin
    TOBLoc := TOBSource.detail[Indice];
    //
    if TOBLoc.GetValue('BLO_QTEDUDETAIL') <> 0 then
    begin
      QTeDDSuite := QteDudetail * TOBLOC.GetValue('BLO_QTEDUDETAIL');
    end else
    begin
      QTeDDSuite := QteDudetail;
    end;
    QteSuite := Qte * TOBLOc.GetValue('BLO_QTEFACT');
    //
    if (TOBLoc.detail.count > 0) then
    begin
      MiseAPlatOuvrageLocal (TOBL,QTeSuite,QteDDSuite,TOBLOC,TOBPLat,DEV);
    end else
    begin
      TOBIns:= TOB.Create ('LIGNEOUV',TOBPlat,-1);
      InsertionChampSupOuv (TOBINS,false);
      TOBIns.AddChampSupValeur ('MONTANTACH',0);
      TOBIns.Dupliquer (TOBLOC,false,true);
      TOBIns.Putvalue ('BLO_QTEFACT',QteSuite/QTeDDSuite);
      CalculMontantHtDevLigOuv (TOBIns,DEV);
      TOBIns.AddChampSupValeur ('MONTANTACHAT',TOBINS.GetValue('BLO_MONTANTPA'));
    end;
  end;
end;

procedure ConstitueOuvragesPlat (TOBpiece,TOBDestination,TOBSource : TOB);
var TOBPlat,TOBL,TOBOUV : TOB;
    Indice : integer;
    IndiceNomen : integer;
    Qte : double;
    DEV : Rdevise;
begin
  //
  DEV.Code := TOBPiece.GetValue('GP_DEVISE');
  GetInfosDevise(DEV);
  //
  TOBPlat := TOB.Create ('PIECE',nil,-1);
  TOBPLat.Dupliquer (TOBPIece,false,true);
  for Indice := 0 to TOBPiece.detail.count -1 do
  begin
    TOBL := TOBPiece.detail[Indice];
    IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');
    Qte := TOBL.GetValue('GL_QTEFACT');
    if (IsOuvrage (TOBL)) and (IndiceNomen<>0) and (IndiceNomen <= TOBSource.detail.count) then
    begin
      TOBPlat.ClearDetail;
      TOBOUV := TOBSource.detail[IndiceNomen-1];
      MiseAPlatOuvrageLocal (TOBL,Qte,1,TOBOUV,TOBPlat,DEV);
      repeat
        TOBPlat.detail[0].ChangeParent (TOBDestination,-1);
      until TOBPlat.detail.count = 0;
    end;
  end;
end;


procedure ConstitueOuvragesDocumentAplat(cledoc : R_CLEDOC; TOBPiece,TOBOUvrages : TOB);
var i : integer;
    TOBLN : TOB;
    DEV:Rdevise;
    valeurs : T_Valeurs;
    TOBOUV : TOB;
begin
  TOBOuv := TOB.Create ('LES OUVRAGES',nil,-1);
  //
  TOBPiece.ClearDetail;
  TOBPiece.InitValeurs;
  //
  TOBOuvrages.ClearDetail;
  TOBOuvrages.InitValeurs;
  //
  BTPLoadLaPiece (cledoc,TOBpiece);
  //
  DEV.Code := TOBPiece.GetValue('GP_DEVISE');
  GetInfosDevise(DEV);
  // Constitution De la tOB Ouvrage
  BTPChargeOuvrages (cledoc,TOBOuv,TOBpiece);
  // Calcul des ouvarges (champs sup)
  for i:= 0 to TOBOuv.detail.count -1 do
  begin
    TOBLN := TOBOuv.detail[I];
    InitTableau (Valeurs);
    CalculeOuvrageDoc (TOBLN,1,1,true,DEV,valeurs,True);
  end;
  // Mise a plat
  ConstitueOuvragesPlat (TOBpiece,TOBOuvrages,TOBOuv);
  //
  TOBOUV.free;
end;

// FONCTIONS GEP REGISTERED

procedure GEPBTPLoadLapiece (Parms : array of variant ; nb : integer) ;
var TOBpiece : TOB;
    Cledoc : R_CLEDOC;
begin
  FillChar(CleDoc,Sizeof(CleDoc),#0) ;
  if Nb = 5 Then
  BEGIN
    TRY
      Cledoc.NaturePiece:= string(Parms[1]);
      Cledoc.Souche  := String (Parms[2]);
      Cledoc.NumeroPiece  := Integer (Parms[3]);
      Cledoc.Indice  := Integer (Parms[4]);
      TOBpiece  := Tob(LongInt(Parms[5]));
    EXCEPT
      PgiBox('Parametres : <Nature>,<Souche>,<Numero>,<Indice>,TOBpiece');
      exit;
    END;
    if TOBPiece=nil then BEGIN PGIBox('TOB Non allouée'); EXIT; END;
    // Init
    BTPLoadLapiece (Cledoc,TOBpiece);
  END ELSE
  BEGIN
    PgiBox('Parametres : <Nature>,<Souche>,<Numero>,<Indice>,TOBpiece');
  END;

end;

procedure GEPBTPChargeOuvrages (Parms : array of variant ; nb : integer) ;
var TOBOuvrages : TOB;
    Cledoc : R_CLEDOC;
begin
  FillChar(CleDoc,Sizeof(CleDoc),#0) ;
  if Nb = 5 Then
  BEGIN
    TRY
      Cledoc.NaturePiece:= string(Parms[1]);
      Cledoc.Souche  := String (Parms[2]);
      Cledoc.NumeroPiece  := Integer (Parms[3]);
      Cledoc.Indice  := Integer (Parms[4]);
      TOBOuvrages  := Tob(LongInt(Parms[5]));
    EXCEPT
      PgiBox('Parametres : <Nature>,<Souche>,<Numero>,<Indice>,TOBOuvrages');
      exit;
    END;
    if TOBOuvrages=nil then BEGIN PGIBox('TOB Non allouée'); EXIT; END;
    // Init
    BTPChargeOuvrages (Cledoc,TOBouvrages);
  END ELSE
  BEGIN
    PgiBox('Parametres : <Nature>,<Souche>,<Numero>,<Indice>,TOBouvrages');
  END;

end;

procedure GEPGetDocumentAPLat (Parms : array of variant ; nb : integer) ;
var TOBOuvrages,TOBpiece : TOB;
    Cledoc : R_CLEDOC;
begin
  FillChar(CleDoc,Sizeof(CleDoc),#0) ;
  if Nb = 6 Then
  BEGIN
    TRY
      Cledoc.NaturePiece:= string(Parms[1]);
      Cledoc.Souche  := String (Parms[2]);
      Cledoc.NumeroPiece  := Integer (Parms[3]);
      Cledoc.Indice  := Integer (Parms[4]);
      TOBpiece  := Tob(LongInt(Parms[5]));
      TOBOuvrages  := Tob(LongInt(Parms[6]));
    EXCEPT
      PgiBox('Parametres : <Nature>,<Souche>,<Numero>,<Indice>,TOBpiece,TOBouvrages');
      exit;
    END;
    if TOBPiece=nil then BEGIN PGIBox('TOB Piece Non allouée'); EXIT; END;
    if TOBOuvrages=nil then BEGIN PGIBox('TOB Ouvrages Non allouée'); EXIT; END;
    ConstitueOuvragesDocumentAplat(Cledoc,TOBpiece,TOBOuvrages);
  END ELSE
  BEGIN
    PgiBox('Parametres : <Nature>,<Souche>,<Numero>,<Indice>,TOBpiece,TOBouvrages');
  END;

end;

procedure GEPLanceFSA (Parms : array of variant ; nb : integer) ;
var filename : string;
    ODExcelFile: TOpenDialog;
begin
	//filename := 'C:\Documents and Settings\All Users\Application Data\POUCHAIN\FSA-POUCHAIN.xls';

  ODExcelFile := TOpenDialog.create(nil) ;
  ODExcelFile.Filter := 'fichiers Excel (*xls;*xlsx)|*.xls;*.xlsx';
  ODExcelFile.InitialDir := TCbpPath.GetCommonAppData+'\POUCHAIN';
  if ODExcelFile.execute then
  begin
    Filename := ODexcelFile.FileName;
  end;
  ODExcelFile.free;

	if filename <> '' then FileExecAndWait (GetExcelPath+'EXCEL.exe "'+filename+'"');
end;

procedure GEPPrepareEditionPouchain (Parms : array of variant ; nb : integer) ;
    procedure VirelesNulls(TOBEtat : TOB);
    var Indice,Ind : integer;
    		TOBD : TOB;
    begin
    	for Indice := 0 to TOBEtat.detail.count -1 do
      begin
      	TOBD := TOBEtat.detail[Indice];
        for Ind := 1000 to 1000+TOBD.ChampsSup.Count -1 do
        begin
          if VarIsNull(TOBD.GetValeur(Ind)) then TOBD.putvaleur(Ind,'');
        end;
      end;
    end;

var TOBEtats,TOBEtat : TOB;
    SQL,Order, Affaire1,Affaire2,Where,Groupe : string;
begin
  if Nb = 2 Then
  BEGIN
		Affaire1 := string(Parms[1]);
    Affaire2 := String (Parms[2]);
  	TOBEtats := TOB.Create('ZETAT', nil, -1);
  	TOBEtats.AddChampSupValeur ('UN CHAMP','');
  	TOBEtat := TOB.Create('ZFEA', TOBEtats, -1);
  	TOBEtat.AddChampSupValeur ('ZZ CHAMP','');
    TRY
    	//
      SQL := RecupSQLComplet ('E', 'GPJ','ZEA', Where,groupe,Order);
      //
      SQL := SQL +' '+Where+ ' AND GL_AFFAIRE1="'+Affaire1+'" AND GL_AFFAIRE2="'+Affaire2+'" '+order;
      TOBEtat.LoadDetailFromSQL  (Sql,false,true);    // recup des lignes
      SQL := StringReplace (SQL,'GL_','BOP_',[rfReplaceAll]);
      SQL := StringReplace (SQL,' LIGNE ',' LIGNEOUVPLAT ',[rfReplaceAll]);
      SQL := StringReplace (SQL,'"MAR","MAR"','"MAR","ARP"',[rfReplaceAll]);
      TOBEtat.LoadDetailFromSQL  (Sql,true,true);    // recup des lignes d'ouvrages à plat

      // traitement frais détaillés de chantier
      SQL := RecupSQLComplet ('E', 'GPJ','ZEA', Where,groupe,Order);
      //
      Where := 'WHERE GP_NATUREPIECEG="FRC" AND GL_TYPELIGNE="ART"';
      Where := Where + ' AND (SELECT P2.GP_LIBREPIECE1 FROM PIECE AS P2 WHERE P2.GP_PIECEFRAIS LIKE "%;FRC;FRC;"+TRIM(CAST(P1.GP_NUMERO AS CHAR))+";0;;") = "X"';
      SQL := SQL +' '+Where+ ' AND GL_AFFAIRE1="'+Affaire1+'" AND GL_AFFAIRE2="'+Affaire2+'" '+order;
      SQL := StringReplace (SQL,'*(GL_COEFFG+1)','',[rfReplaceAll]);
      SQL := StringReplace (SQL,'GL_MONTANTPA','GL_MONTANTHT',[rfReplaceAll]);

      TOBEtat.LoadDetailFromSQL  (Sql,True,true);    // recup des lignes
      SQL := StringReplace (SQL,'GL_','BOP_',[rfReplaceAll]);
      SQL := StringReplace (SQL,' LIGNE ',' LIGNEOUVPLAT ',[rfReplaceAll]);
      SQL := StringReplace (SQL,'"MAR","MAR"','"MAR","ARP"',[rfReplaceAll]);
      TOBEtat.LoadDetailFromSQL  (Sql,true,true);    // recup des lignes d'ouvrages à plat
      //

      if TOBEtat.detail.count > 0 then
      begin
      	VirelesNulls(TOBEtat);
      	LanceEtatTob('E','GPJ','ZEA',TOBEtats,true,false,false,nil,'','Etat FEA',false);
      end else
      begin
        PGIBox(TraduireMemoire('Aucune ligne à traiter'));
      end;
    FINALLY
    	TOBEtat.free;
    END;
  END ELSE
  BEGIN
    PgiBox('Parametres : <CODEAFFAIRE1>,<CODEAFFAIRE2>');
  END;
end;

// SPECIF LSE

procedure AjouteFactureExport (cledoc : r_cledoc; TOBPiecesExp : TOB) ;
var Sql : string;
    QQ : TQuery;
    TOBPiece : TOB;
begin
  Sql := 'SELECT GP_NATUREPIECEG,GP_SOUCHE,GP_NUMERO,GP_INDICEG,GP_TIERSFACTURE,'+
         'GP_DATEPIECE,GP_MODEREGLE,GPE_MODEPAIE,GPE_DATEECHE,T_AUXILIAIRE, YTC_TEXTELIBRE2 '+
         'FROM PIECE '+
         'LEFT JOIN PIEDECHE ON GPE_NATUREPIECEG=GP_NATUREPIECEG AND GPE_SOUCHE=GP_SOUCHE AND '+
         'GPE_NUMERO=GP_NUMERO AND GPE_INDICEG=GP_INDICEG AND GPE_NUMECHE=1 '+
         'LEFT JOIN TIERS ON T_TIERS=GP_TIERSFACTURE AND T_NATUREAUXI="CLI" '+
         'LEFT JOIN TIERSCOMPL ON YTC_TIERS=GP_TIERSFACTURE ' +
         'WHERE '+ WherePiece(cledoc,ttdPiece,true);
  TRY
    QQ := OpenSQL (Sql,true,1,'',true);
    if not QQ.Eof then
    begin
      TOBPiece := TOB.Create ('PIECE',TOBPiecesExp,-1);
      TOBPiece.SelectDB('',QQ);
    end;
  FINALLY
    Ferme (QQ);
  END;
end;

//chargement de la TOB des tiers compléments
procedure LoadLesTiersCompl(TOBPiece, TobTiersCompl : TOB);
var StSQl : string;
    QQ    : TQuery;
begin

  StSQl := 'SELECT YTC_TEXTELIBRE2 FROM TIERSCOMPL WHERE YTC_TIERS="' + TOBPiece.GetString('T_TIERS') + '"';
  QQ := OpenSQL(StSQl, True);

  if not QQ.Eof then
  begin
    TobTiersCompl.SelectDB('',QQ);
  end;

  Ferme(QQ);

end;

function ConstitueEntete (TobPiece,TOBAdresses : TOB) : string;
begin
	result := TrimLeft(trimright(IntToStr(TOBPiece.GetValue('GP_NUMERO'))))+
  					TabLse +
            //TrimLeft(trimRight(TOBPiece.GetValue('T_AUXILIAIRE'))) +
            TrimLeft(trimRight(TOBPiece.GetString('YTC_TEXTELIBRE2'))) +
  					TabLse +
            DateToStr(TOBPiece.GetValue('GP_DATEPIECE')) +
  					TabLse +
            TrimLeft(TrimRight(TobAdresses.detail[1].GetValue('GPA_LIBELLE')))+
  					TabLse +
            TrimLeft(TrimRight(TobAdresses.detail[1].GetValue('GPA_ADRESSE1')))+
  					TabLse +
            TrimLeft(TrimRight(TobAdresses.detail[1].GetValue('GPA_ADRESSE2')))+
  					TabLse +
            TrimLeft(TrimRight(TobAdresses.detail[1].GetValue('GPA_ADRESSE3')))+
  					TabLse +
            TrimLeft(TrimRight(TobAdresses.detail[1].GetValue('GPA_CODEPOSTAL')))+
  					TabLse +
            TrimLeft(TrimRight(TobAdresses.detail[1].GetValue('GPA_VILLE')))+
  					TabLse +
            TrimLeft(TrimRight(TobAdresses.detail[0].GetValue('GPA_LIBELLE')))+
  					TabLse +
            TrimLeft(TrimRight(TobAdresses.detail[0].GetValue('GPA_ADRESSE1')))+
  					TabLse +
            TrimLeft(TrimRight(TobAdresses.detail[0].GetValue('GPA_ADRESSE2')))+
  					TabLse +
            TrimLeft(TrimRight(TobAdresses.detail[0].GetValue('GPA_ADRESSE3')))+
  					TabLse +
            TrimLeft(TrimRight(TobAdresses.detail[0].GetValue('GPA_CODEPOSTAL')))+
  					TabLse +
            TrimLeft(TrimRight(TobAdresses.detail[0].GetValue('GPA_VILLE')))+
  					TabLse +
            trimLeft(trimRight(TOBPiece.GetValue('GPE_MODEPAIE'))) +
  					TabLse +
            trimLeft(trimRight(TOBPiece.GetValue('GP_MODEREGLE'))) +
  					TabLse +
            trimLeft(trimRight(TOBPiece.GetValue('GPE_DATEECHE')));
  //FileWrite (F,Buffer,length(Buffer));
//  WriteLn(F,buffer);
end;

function  ConstitueLineDetail(TOBLigne : TOB) : string;
begin
	result := TrimLeft(trimright(IntToStr(TOBLigne.GetValue('GL_NUMERO'))))+
  					TabLse +
            TrimLeft(trimRight(TOBLigne.GetValue('GA_CHARLIBRE1'))) +
  					TabLse +
            TrimLeft(trimRight(TOBLigne.GetValue('GL_LIBELLE'))) +
  					TabLse +
            TrimLeft(trimRight(STRFPOINT(TOBLigne.GetValue('GL_QTEFACT')))) +
  					TabLse +
            TrimLeft(trimRight(STRFPOINT(TOBLigne.GetValue('GL_PUHTDEV')))) +
            TabLse +
            TrimLeft(trimRight(STRFPOINT(TOBLigne.GetValue('GL_DPA'))));
  //FileWrite (FD, Buffer,length(Buffer));
end;

procedure  ConstituelesLignes (TOBpiece,TOBLignes : TOB);
var Sql,refPiece : string;
    Cledoc : R_CLEDOC;
begin
	RefPiece := EncodeRefPiece(TOBPiece,0);
  DecodeRefPiece(RefPiece,Cledoc);
  Sql := 'SELECT GL_NUMERO,GL_CODEARTICLE,GL_LIBELLE,GL_QTEFACT,GL_PUHTDEV, GL_DPA,GA_CHARLIBRE1 '+
         'FROM LIGNE '+
         'LEFT JOIN ARTICLE ON GA_ARTICLE=GL_ARTICLE '+
         'WHERE '+ WherePiece(cledoc,ttdLigne,false)+' '+
         'ORDER BY GL_NUMLIGNE';
	TOBLignes.LoadDetailFromSQL(Sql,false);
end;

procedure GEPExportefactures (Parms : array of variant ; nb : integer) ;
var Indice,IInd : integer;
		EmplacementExport,NomEnt,NomDet : string;
    NaturePiece : string;
    EtatAffaire : string;
    EtatVisa    : string;
    TOBPiece,TOBAdresses,TOBPiecesExp,TOBLignes,TOBArticlesP : TOB;
    TobTiersCompl : TOB;
    F,FD : textfile;
    cledoc : r_cledoc;
    Q,QQ : TQuery;
    XX : TForm;
    buffer,SQL : string;
    LesErreurs : TlistErrorExport;
    fcledoc : r_cledoc;
begin

  XX:=TForm(Longint(Parms[0]));
  if (TFMul(XX).FListe=nil) then exit;
  EmplacementExport:= string(Parms[1]);
  //
  NaturePiece := String(Parms[2]);
  EtatAffaire := String(Parms[3]);
  EtatVisa    := String(Parms[4]);
  //
  if EmplacementExport = '' then BEGIN PGIError ('Renseignez le répertoire d''export'); exit; END;

  if (not TFMul(XX).Fliste.AllSelected) and (TFMul(XX).Fliste.NbSelected=0) then
  begin
    PgiInfo ('Aucun document n''est sélectionné');
    Exit;
  end;

  if (PGIAsk ('Désirez-vous réellement exporter vos documents', 'Export')<>mrYes) then exit;

  LesErreurs := TlistErrorExport.Create;

  if NaturePiece = 'DBT' then
  begin
    NomEnt := IncludeTrailingBackslash(EmplacementExport)+'Commandes.txt';
    NomDet := IncludeTrailingBackslash(EmplacementExport)+'LignesCdes.txt';
  end
  else
  begin
    NomEnt := IncludeTrailingBackslash(EmplacementExport)+'Documents.txt';
    NomDet := IncludeTrailingBackslash(EmplacementExport)+'Lignes.txt';
  end;

  if not DirectoryExists(EmplacementExport) then
  begin
    PGIError ('Répertoire inexistant');
    exit;
  end;

  TOBPiecesExp  := TOB.Create ('LES PIECES',nil,-1);
  TOBAdresses   := TOB.Create ('LES ADRESSES', nil,-1);
  Toblignes     := TOB.Create ('LES LIGNES', nil,-1);
  TobTiersCompl := TOB.Create ('LES COMPLEMENTS', nil, -1);
  TOBArticlesP  := TOB.Create ('LES ARTICLES PIECES',nil,-1);

  if FileExists(NomEnt) then DeleteFile(PAnsiChar(NomEnt));
  if FileExists(Nomdet) then DeleteFile(PansiChar(NomDet));

	AssignFile(F,NomEnt); SetLineBreakStyle(F,tlbsCRLF);
	AssignFile(FD,NomDet); SetLineBreakStyle(FD,tlbsCRLF);

  rewrite(F);
  rewrite(FD);

  TRY
    if TFMul(XX).Fliste.AllSelected then
    BEGIN
      Q:=TFMul(XX).Q;
      Q.First;
      while Not Q.EOF do
      BEGIN
        fillchar(cledoc,Sizeof(CleDoc),#0);
        cledoc.NaturePiece:=Q.FindField('GP_NATUREPIECEG').AsString;
        cledoc.Souche :=Q.FindField('GP_SOUCHE').AsString;
        Cledoc.NumeroPiece :=Q.FindField('GP_NUMERO').AsInteger;
        Cledoc.Indice  :=Q.FindField('GP_INDICEG').AsInteger;
        AjouteFactureExport (cledoc,TOBPiecesExp);
        Q.NEXT;
      END;
      TFMul(XX).Fliste.AllSelected:=False;
    END ELSE
    BEGIN
      for indice:=0 to TFMul(XX).Fliste.nbSelected-1 do
      begin
        TFMul(XX).Fliste.GotoLeBookmark(indice);
        fillchar(cledoc,Sizeof(CleDoc),#0);
        cledoc.NaturePiece:=TFMul(XX).Fliste.datasource.dataset.FindField('GP_NATUREPIECEG').AsString;
        cledoc.Souche:=TFMul(XX).Fliste.datasource.dataset.FindField('GP_SOUCHE').AsString;
        Cledoc.NumeroPiece:=TFMul(XX).Fliste.datasource.dataset.FindField('GP_NUMERO').Asinteger;
        Cledoc.Indice  := TFMul(XX).Fliste.datasource.dataset.FindField('GP_INDICEG').AsInteger;
        AjouteFactureExport (cledoc,TOBPiecesExp);
      end;
    END;

    // Traitement lignes détail
    if TOBPiecesExp.detail.count = 0 then exit;
    // phase de vérification
    for Indice := 0 to TOBPiecesExp.detail.count -1 do
    begin
      TOBPiece := TOBPiecesExp.detail[Indice];
      fcledoc := TOB2CleDoc  (TOBPiece);
      if TobPiece.GetString('YTC_TEXTELIBRE2') = '' then
      begin
        (*
        PgiInfo  ('La zone texte libre 2 (Référence client L.S.E) du tiers ' + TOBPiece.GetString('GP_TIERSFACTURE') +
                  ' est vide.#13#10 Le devis N°' + TOBPIece.GetString('GP_NUMERO') +
                  ' ne peut pas être exporté !');
        *)
      	QQ := OpenSQL('SELECT T_AUXILIAIRE,T_NATUREAUXI FROM TIERS WHERE T_TIERS="' + TOBPiece.GetString('GP_TIERSFACTURE') + '"', True,-1, '', True);
        if not QQ.eof then
        begin
        	LesErreurs.Add(TteTiers ,TOBPiece.GetString('GP_TIERSFACTURE'),QQ.findField('T_AUXILIAIRE').asstring,QQ.findField('T_NATUREAUXI').asstring,'','','');
        end;
      	ferme (QQ);
      end;
      SQL := 'SELECT GA_ARTICLE,GA_CODEARTICLE,GA_TYPEARTICLE,GA_CHARLIBRE1 FROM ARTICLE WHERE GA_ARTICLE IN '+
      			  '(SELECT DISTINCT GL_ARTICLE FROM LIGNE WHERE '+WherePiece(fcledoc,ttdLigne,false)+
              ' AND GL_TYPELIGNE="ART" )';

      TOBArticlesP.LoadDetailDBFromSQL  ('ARTICLE',Sql,false);
      for IInd := 0 to TOBArticlesP.Detail.count -1 do
      begin
        if TOBArticlesP.detail[iind].GetString('GA_CHARLIBRE1') = '' then
        begin
        	LesErreurs.Add(TteArticle ,'','','',TOBArticlesP.detail[iind].GetString('GA_ARTICLE'),
          							TOBArticlesP.detail[iind].GetString('GA_CODEARTICLE'),
                        TOBArticlesP.detail[iind].GetString('GA_TYPEARTICLE'));
        end;
      end;
    end;
    TOBArticlesP.ClearDetail;

    if LesErreurs.InErrorCount > 0 then
    begin
      TraitelesErreurs (LesErreurs);
    end;

    if LesErreurs.InErrorCount > 0 then exit;


    for Indice := 0 to TOBPiecesExp.detail.count -1 do
    begin
      TOBPiece := TOBPiecesExp.detail[Indice];
      LoadLesAdresses(TOBPiece,TOBAdresses);
      buffer := ConstitueEntete (TobPiece,TOBAdresses);
      writeln (F,buffer);
      ConstituelesLignes (TOBpiece,TOBLignes);
      //
      for IInd := 0 to TOBLignes.detail.count -1 do
      begin
        Buffer := ConstitueLineDetail (TOBLignes.detail[IInd]);
      	writeln (FD,buffer);
      end;
      TOBPiece.PutValue('GP_ETATVISA', 'VIS');
      TOBPiece.UpdateDB;
      TOBAdresses.clearDetail;
      TobTiersCompl.ClearDetail;
      TOBLignes.cleardetail;
    end;
  FINALLY
    CloseFile(F);
    CloseFile(FD);
    TOBPiecesExp.free;
    TOBAdresses.Free;
    FreeAndNil(TobTiersCompl);
    Toblignes.free;
    TOBArticlesP.free;
    LesErreurs.Free;
    TFMul(XX).Fliste.ClearSelected;
    TFMUL(XX).BChercheClick(Nil);
  END;
end;

procedure GEPLanceExcelLSE (Parms : array of variant ; nb : integer) ;
var filename,repertLoc : string;
		XX : Tform;
begin
  XX:=TForm(Longint(Parms[0]));
  if (TFMul(XX).FListe=nil) then exit;
	repertLoc := TCBPPath.GetPersonal;
	FileName := IncludeTrailingBackslash(GetParamSocSecur('SO_DIREXPORTSXLS',RepertLoc))+
  						string(Parms[1])+'.xlsm';
	if FileExists(FileName) then
  begin
		FileExecAndWait (GetExcelPath+'EXCEL.exe "'+filename+'"');
  end;
end;

procedure RegisterGepMethods();
begin
 // GEP
 RegisterAglProc( 'BTPGetDocumentAPLat',True,6,GEPGetDocumentAPLat);
 RegisterAglProc( 'BTPLoadPiece',True,5,GEPBTPLoadLapiece);
 RegisterAglProc( 'BTPLoadOuvrages',True,5,GEPBTPChargeOuvrages);
 // SPECIF POUCHAIN
 RegisterAglProc( 'BTPPrepareEditionPouchain',True,2,GEPPrepareEditionPouchain);
 RegisterAglProc( 'BTPLanceFSA',True,0,GEPLanceFSA);
 RegisterAglProc( 'GEPLanceExcelLSE',True,1,GEPLanceExcelLSE);
 //
 // SPECIF LSE
 RegisterAglProc( 'BTPExportefactures',True,2,GEPExportefactures);
 //
end;

Initialization
RegisterGepMethods();
end.
