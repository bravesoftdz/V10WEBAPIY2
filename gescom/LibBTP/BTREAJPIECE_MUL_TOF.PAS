{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 23/05/2017
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTREAJPIECE_MUL ()
Mots clefs ... : TOF;BTREAJPIECE_MUL
*****************************************************************}
Unit BTREAJPIECE_MUL_TOF ;

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
     M3Fp,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     uTofAfBaseCodeAffaire,
     HTB97,
     uTob,
     EntGc,
     RECALCPIECE_RAP_TOF,
     FactTob,
     Ent1,uEntCommun,UtilTOBPiece,UtilsRapport,SaisUtil ;

Type
  TOF_BTREAJPIECE_MUL = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    DEV : Rdevise;
  	TOBPieces : TOB;
    TheSituations : TOB;
    OneRapport : TGestionRapport;
    procedure AddPieceAtraiter(Cledoc: r_cledoc; lesPieces : TOB);
    procedure Reajusteanterieurs;
    procedure Controlechamp(Champs, valeur: string);
    procedure ReajusteCesPieces (TOBPieces : TOB);
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4,Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit); 
    function TraiteRecupAvancPrec (TOBPiece : TOB) : Boolean;
  end ;

Implementation
uses FactureBtp;

procedure TOF_BTREAJPIECE_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTREAJPIECE_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTREAJPIECE_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTREAJPIECE_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTREAJPIECE_MUL.OnArgument (S : String ) ;
var CC : THValComboBox;
    critere,champ,valeur : string;
    i : Integer;
begin
  Inherited ;

  TOBPieces := TOB.create ('LES PIECES',nil,-1);
  TheSituations := TOB.Create ('LES SITUATIONS',nil,-1);
  //
  CC:=THValComboBox(GetControl('GP_ETABLISSEMENT')) ;
  if CC<>Nil then PositionneEtabUser(CC) ;
  //
  //Gestion Restriction Domaine
  CC:=THValComboBox(GetControl('GP_DOMAINE')) ;
  if CC<>Nil then PositionneDomaineUser(CC) ;

  While (Critere <> '') do
  BEGIN
    i:=pos(':',Critere);
    if i = 0 then i:=pos('=',Critere);
    if i <> 0 then
       begin
       Champ:=copy(Critere,1,i-1);
       Valeur:=Copy (Critere,i+1,length(Critere)-i);
       end
    else
       Champ := Critere;
    Controlechamp(Champ, Valeur);
    Critere:=(Trim(ReadTokenSt(S)));
  END;
  OneRapport       := TGestionRapport.Create(TForm(Ecran));
  with OneRapport do
  begin
    Titre   := 'Recalcul des cumulés antérieurs';
    Affiche := true;
    Close   := True;
    Sauve   := False;
    Print   := False;
  end;
end ;

procedure TOF_BTREAJPIECE_MUL.OnClose ;
begin
  TheSituations.free;
	OneRapport.Free;
  FreeAndNil(TOBPieces);
  Inherited ;
end ;

procedure TOF_BTREAJPIECE_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTREAJPIECE_MUL.OnCancel () ;
begin
  Inherited ;
end ;


procedure TOF_BTREAJPIECE_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1,
  Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers,
  Tiers_: THEdit); 
begin
  inherited;
  Aff:=THEdit(GetControl('GP_AFFAIRE'))   ;
  Aff0:=THEdit(GetControl('GP_AFFAIRE0'));
  Aff1:=THEdit(GetControl('GP_AFFAIRE1')) ;
  Aff2:=THEdit(GetControl('GP_AFFAIRE2')) ;
  Aff3:=THEdit(GetControl('GP_AFFAIRE3')) ;
  Aff4:=THEdit(GetControl('GP_AVENANT'))  ;
  Tiers:=THEdit(GetControl('GP_TIERS'))   ;
end;

procedure TOF_BTREAJPIECE_MUL.Reajusteanterieurs;
var Q : Tquery;
		F : TFmul;
    i : integer;
    Cledoc : r_cledoc;
begin
  F := TFmul(ecran);
  TheSituations.ClearDetail;
  TOBPieces.ClearDetail;
  TOBPieces.InitValeurs;

  if (PGIAsk ('Désirez-vous réajuster les situations sélectionnées ?', ecran.Caption)<>mrYes) then exit;

  TRY
    if TFMul(F).Fliste.AllSelected then
    BEGIN
      Q:=TFmul(F).Q;
      Q.First;
      while Not Q.EOF do
      BEGIN
        FillChar(CleDoc, Sizeof(CleDoc), #0);
        cledoc.NaturePiece := Q.FindField('GP_NATUREPIECEG').AsString;
        Cledoc.Souche      := Q.FindField('GP_SOUCHE').AsString;
        cledoc.NumeroPiece := Q.FindField('GP_NUMERO').AsInteger;
        cledoc.DatePiece   := Q.FindField('GP_DATEPIECE').AsDateTime;
        cledoc.Indice      := Q.FindField('GP_INDICEG').AsInteger;
        AddPieceAtraiter (Cledoc,TOBPieces);
        Q.NEXT;
      END;
      TFMul(F).Fliste.AllSelected:=False;
    END ELSE
    BEGIN
      for i:=0 to TFMul(F).Fliste.nbSelected-1 do
      begin
        TFMul(F).Fliste.GotoLeBookmark(i);
        FillChar(CleDoc, Sizeof(CleDoc), #0);
        cledoc.NaturePiece :=TFMul(F).Fliste.datasource.dataset.FindField('GP_NATUREPIECEG').AsString;
        Cledoc.Souche      :=TFMul(F).Fliste.datasource.dataset.FindField('GP_SOUCHE').AsString;
        cledoc.NumeroPiece :=TFMul(F).Fliste.datasource.dataset.FindField('GP_NUMERO').AsInteger;
        cledoc.DatePiece   :=TFMul(F).Fliste.datasource.dataset.FindField('GP_DATEPIECE').AsDateTime;
        cledoc.Indice      :=TFMul(F).Fliste.datasource.dataset.FindField('GP_INDICEG').AsInteger;
        AddPieceAtraiter (Cledoc,TOBPieces);
      end;
    END;
  finally
    if TOBPieces.detail.count > 0 then ReajusteCesPieces (TOBPieces);
  	TOBPieces.ClearDetail;
  end;
end;

procedure TOF_BTREAJPIECE_MUL.AddPieceAtraiter(Cledoc: r_cledoc; LesPieces : TOB);
var TOBPiece : TOB;
		QQ : Tquery;
    SQl : string;
begin
	TOBPIece := TOB.Create ('PIECE',LesPieces,-1);
  SQL := 'SELECT '+
         'GP_NATUREPIECEG,GP_SOUCHE,GP_NUMERO,GP_INDICEG,GP_AFFAIREDEVIS,BST_NUMEROSIT '+
         'FROM PIECE '+
         'LEFT JOIN BSITUATIONS ON BST_NATUREPIECE=GP_NATUREPIECEG AND BST_SOUCHE=GP_SOUCHE AND BST_SSAFFAIRE=GP_AFFAIREDEVIS AND BST_NUMEROFAC=GP_NUMERO '+
         'WHERE '+WherePiece (cledoc,ttdPiece,true) +' AND BST_VIVANTE="X"';
  QQ := OpenSql (SQL,True);
  TOBPiece.SelectDB ('',QQ);
  Ferme (QQ);
end;

procedure AglBTReajusteanterieur (parms:array of variant; nb: integer ) ;
var  F : TForm ;
     LaTof : TOF;
begin

  F:=TForm(Longint(Parms[0])) ;

  if (F is TFMul) then Latof:=TFMul(F).Latof else exit;

  TOF_BTREAJPIECE_MUL(LaTof).Reajusteanterieurs
end;

procedure TOF_BTREAJPIECE_MUL.ReajusteCesPieces(TOBPieces: TOB);
var II : Integer;
begin
  OneRapport.InitRapport;
  TOBPieces.detail.Sort('GP_AFFAIREDEVIS;BST_NUMEROSIT');
  for ii := 0 to TOBPieces.detail.count -1 do
  begin
    if TOBPieces.detail[II].GetInteger('BST_NUMEROSIT') > 1 then
    begin
      if TraiteRecupAvancPrec(TOBPieces.detail[II]) then
      begin
        OneRapport.SauveLigMemo('Document '+TOBPieces.detail[II].GetString('GP_NUMERO')+ ' -> OK ');
      end;
    end;
  end;
  OneRapport.AfficheRapport;
end;


procedure TOF_BTREAJPIECE_MUL.Controlechamp(Champs, valeur: string);
begin
//
end;

function TOF_BTREAJPIECE_MUL.TraiteRecupAvancPrec(TOBPiece: TOB): Boolean;

  function LoadPiece (TOBPiece : TOB) : Boolean;
  var QQ : TQuery;
      SQl : string;
  begin
    Result := false;
    SQL := 'SELECT LIGNEFAC.*,GL_TYPEARTICLE,GL_ARTICLE,GL_LIBELLE,GL_NUMLIGNE AS NUMLIG,GL_NUMORDRE,GL_PIECEPRECEDENTE,GL_PIECEORIGINE,GL_QTEPREVAVANC,GL_MONTANTHTDEV '+
           'FROM LIGNEFAC '+
           'LEFT JOIN LIGNE ON '+
           'BLF_NATUREPIECEG=GL_NATUREPIECEG AND BLF_SOUCHE=GL_SOUCHE AND BLF_NUMERO=GL_NUMERO AND BLF_NUMORDRE = GL_NUMORDRE '+
           'WHERE '+
           'BLF_NATUREPIECEG="'+TOBPiece.GetString('GP_NATUREPIECEG')+'" AND '+
           'BLF_SOUCHE="'+TOBPiece.GetString('GP_SOUCHE')+'" AND '+
           'BLF_NUMERO="'+TOBPiece.GetString('GP_NUMERO')+'" AND '+
           'GL_ARTICLE <> '' AND '+
           'BLF_UNIQUEBLO=0 '+
           'ORDER BY GL_NUMLIGNE';
    QQ := OpenSQL(SQl,True,-1,'',true);
    if Not QQ.eof then
    begin
      TOBPiece.LoadDetailDB('LIGNEFAC','','',QQ,false);
      Result := True;
    end;
    ferme (QQ);
  end;

  function LoadLesOuvrages (TOBPiece : TOB) : Boolean;
  var QQ : TQuery;
      SQl : string;
      LastInd,II,JJ : Integer;
      TOBO : TOB;
      okok : Boolean;
      TOBOUvrages : TOB;
  begin
    TOBOUvrages := TOB.Create ('LES OUVRAGES',nil,-1);
    Result := true;
    SQL := 'SELECT LIGNEFAC.*,BLO_TYPEARTICLE,BLO_ARTICLE,BLO_LIBELLE,BLO_NUMLIGNE AS NUMLIG,BLO_N1,BLO_UNIQUEBLO,BLO_PIECEPRECEDENTE '+
           'FROM LIGNEFAC '+
           'LEFT JOIN LIGNEOUV ON '+
           'BLF_NATUREPIECEG=BLO_NATUREPIECEG AND BLF_SOUCHE=BLO_SOUCHE AND BLF_NUMERO=BLO_NUMERO AND BLF_UNIQUEBLO = BLO_UNIQUEBLO '+
           'WHERE '+
           'BLF_NATUREPIECEG="'+TOBPiece.GetString('GP_NATUREPIECEG')+'" AND '+
           'BLF_SOUCHE="'+TOBPiece.GetString('GP_SOUCHE')+'" AND '+
           'BLF_NUMERO="'+TOBPiece.GetString('GP_NUMERO')+'" AND '+
           'BLF_UNIQUEBLO<>0 AND BLO_N2=0 '+
           'ORDER BY BLO_NUMLIGNE';
    TRY
      TRY
        QQ := OpenSQL(SQl,True,-1,'',true);
        if Not QQ.eof then
        begin
          TOBOuvrages.LoadDetailDB('LIGNEFAC','','',QQ,false);
        end;
        
        if TOBOUvrages.Detail.count > 0 then
        begin
          Result := false;
          LastInd := 0;
          II := 0;
          repeat
            TOBO := TOBOUvrages.Detail[II];
            okok := false;
            for JJ := LastInd to TOBPiece.Detail.count - 1 do
            begin
              if TOBPiece.detail[JJ].GetInteger('NUMLIG')=TOBO.GetInteger('NUMLIG') then
              begin
                okok := True;
                Break;
              end;
            end;
            if not Okok then begin Inc(II); Continue; end;
            TOBO.ChangeParent(TOBPiece.detail[JJ],-1);
            LastInd := JJ;
          until II >= TOBOUvrages.Detail.count;
        end;
        Result := True;
      FINALLY
        ferme (QQ);
      end;
    except
      Exit;
    End;
  end;

  function FindSitPrec (TOBPiece : TOB) : TOB;
  var OTOB : TOB;
  begin
    result := nil;
    OTOB := TheSituations.FindFirst(['BST_NUMEROFAC'],[TOBPiece.getInteger('GP_NUMERO')],true);
    if OTOB.GetIndex = 0 then exit;
    result := TheSituations.detail[OTOB.GetIndex-1];
  end;

  function FindLigneSitPrec (TOBSitPre : TOB; PiecePrec : string; TOBAvancPre : TOB) : boolean;
  var SQL : string;
      QQ : TQuery;
  begin
    result:= false;
    SQL := 'SELECT LIGNEFAC.*,GL_TYPEARTICLE,GL_ARTICLE,GL_LIBELLE,GL_NUMLIGNE AS NUMLIG,GL_NUMORDRE,GL_PIECEPRECEDENTE,GL_PIECEORIGINE,GL_QTEPREVAVANC '+
           'FROM LIGNEFAC '+
           'LEFT JOIN LIGNE ON '+
           'BLF_NATUREPIECEG=GL_NATUREPIECEG AND BLF_SOUCHE=GL_SOUCHE AND BLF_NUMERO=GL_NUMERO AND BLF_NUMORDRE = GL_NUMORDRE '+
           'WHERE '+
           'BLF_NATUREPIECEG="'+TOBSitPre.GetString('BST_NATUREPIECE')+'" AND '+
           'BLF_SOUCHE="'+TOBSitPre.GetString('BST_SOUCHE')+'" AND '+
           'BLF_NUMERO="'+TOBSitPre.GetString('BST_NUMEROFAC')+'" AND '+
           'GL_PIECEPRECEDENTE="'+PiecePrec+'"';
    QQ := OpenSQL(SQL,True,1,'',true);
    if not QQ.eof then
    begin
      TOBAvancPre.SelectDB('',QQ);
      Result := True;
    end;
    ferme (QQ);
  end;

  procedure CalculeAvanc (TheAvancPrec : TOB);
  var II : integer;
      TOBLO : TOB;
      QteDet,Avancprec,{QteDejaFact,}MtDetDev{,MtDejafact} : double;
  begin
    for II := 0 to TheAvancPrec.detail.count -1 do
    begin
      TOBLO := TheAvancPrec.detail[II];
      if TOBLO.GetString('BLF_NATUREPIECEG')='' then
      begin
        // cas de figure ou les sous détails ont disparus
        QteDet := arrondi(TOBLO.getValue('BLO_QTEFACT')*TheAvancPrec.getvalue('GL_QTEPREVAVANC') /TOBLO.getValue('BLO_QTEDUDETAIL'),V_PGI.okdecQ);
        if TheAvancPrec.Getvalue('BLF_MTMARCHE') <> 0 then
        begin
          avancprec := TheAvancPrec.Getvalue('BLF_MTDEJAFACT')/TheAvancPrec.Getvalue('BLF_MTMARCHE'); // ratio pour deja facturé
        end else
        begin
          Avancprec := 0;
        end;
//        QteDejaFact := Arrondi(QteDet * Avancprec,V_PGI.okdecQ);
        MtDetDev := arrondi(QteDet*TOBLO.getValue('BLO_PUHTDEV'),DEV.Decimale);
//        MtDejafact := arrondi(QteDejafact*TOBLO.getValue('BLO_PUHTDEV'),DEV.Decimale);
        //
        TOBLO.putValue('BLF_POURCENTAVANC',TheAvancPrec.Getvalue('BLF_POURCENTAVANC'));
        //
//        QteDejaFact := arrondi(TOBLO.getValue('BLO_QTEFACT')*TheAvancPrec.getvalue('GL_QTEPREVAVANC') /TOBLO.getValue('BLO_QTEDUDETAIL'),V_PGI.okdecQ);
        //
        TOBLO.putValue('BLF_MTMARCHE',MtDetDev);
        TOBLO.putValue('BLF_MTCUMULEFACT',arrondi(MtDetDev*TheAvancPrec.Getvalue('BLF_POURCENTAVANC')/100,DEV.Decimale ));
        TOBLO.putValue('BLF_MTDEJAFACT',arrondi(MtDetDev*avancprec,DEV.decimale));
        TOBLO.putValue('BLF_MTSITUATION',arrondi(TOBLO.getValue('BLF_MTCUMULEFACT')-TOBLO.getValue('BLF_MTDEJAFACT'),DEV.Decimale));
        TOBLO.putValue('BLF_QTEMARCHE',QteDet);
        TOBLO.putValue('BLF_QTECUMULEFACT',arrondi(QteDet*TOBLO.Getvalue('BLF_POURCENTAVANC')/100,V_PGI.okdecQ));
        TOBLO.putValue('BLF_QTEDEJAFACT',arrondi(QteDet*avancprec,V_PGI.okdecQ));
        TOBLO.putValue('BLF_QTESITUATION',arrondi(TOBLO.GetValue('BLF_QTECUMULEFACT')-TOBLO.GetValue('BLF_QTEDEJAFACT'),V_PGI.okdecQ));
        TOBLO.putValue('BLF_NATUREPIECEG',TheAvancPrec.Getvalue('BLF_NATUREPIECEG'));
        TOBLO.putValue('BLF_SOUCHE',TheAvancPrec.Getvalue('BLF_SOUCHE'));
        TOBLO.putValue('BLF_NUMERO',TheAvancPrec.Getvalue('BLF_NUMERO'));
        TOBLO.putValue('BLF_INDICEG',TheAvancPrec.Getvalue('BLF_INDICEG'));
        TOBLO.putValue('BLF_NUMORDRE',0);
        TOBLO.putValue('BLF_UNIQUEBLO',TOBLO.Getvalue('BLO_UNIQUEBLO'));
        TOBLO.putValue('BLF_NATURETRAVAIL',TOBLO.Getvalue('BLO_NATURETRAVAIL'));
        TOBLO.putValue('BLF_FOURNISSEUR',TOBLO.Getvalue('BLO_FOURNISSEUR'));
      end;
    end;
  end;

  function FindSousDetail (TheAvancLPrec : TOB; PiecePrecedente : string) : TOB;
  begin
    Result := TheAvancLPrec.FindFirst(['BLO_PIECEPRECEDENTE'],[PiecePrecedente],true);
  end;


  procedure AppliqueAvancLignePrec (TheAvancLPrec,OTOB : TOB);
  var TOBPDOUV,TOBLO : TOB;
      II : integer;
      Pavanc : double;
  begin
    // on récupère les choses bizares arrivant chez LEFUNTEUN
    OTOB.SetDouble('BLF_MTSITUATION',OTOB.geTDouble('GL_MONTANTHTDEV'));
    // on garde l'avancement en montant et qte, on recupere l'avancement précédent et on recalcule le pourcentage d'avancement.
    OTOB.SetDouble('BLF_MTCUMULEFACT',TheAvancLPrec.GetDouble('BLF_MTCUMULEFACT')+OTOB.geTDouble('BLF_MTSITUATION'));
    OTOB.SetDouble('BLF_MTDEJAFACT',TheAvancLPrec.GetDouble('BLF_MTCUMULEFACT'));
    //
    OTOB.SetDouble('BLF_QTECUMULEFACT',TheAvancLPrec.GetDouble('BLF_QTECUMULEFACT')+OTOB.geTDouble('BLF_QTESITUATION'));
    OTOB.SetDouble('BLF_QTEDEJAFACT',TheAvancLPrec.GetDouble('BLF_QTECUMULEFACT'));
    //
    if OTOB.GetDouble('BLF_MTMARCHE') <> 0 then PAvanc := Arrondi( (OTOB.GetDouble('BLF_MTCUMULEFACT')/OTOB.GetDouble('BLF_MTMARCHE'))*100,2)
                                           else Pavanc := 0;
    OTOB.SetDouble('BLF_POURCENTAVANC',Pavanc);
    //
    if isOuvrage(OTOB) then
    begin
      if OTOB.Detail.count > 0 then
      begin
        for II := 0 TO OTOB.detail.count -1 do
        begin
          TOBLO := OTOB.detail[II];
          TOBLO.SetDouble('BLF_MTDEJAFACT',0);
          TOBLO.SetDouble('BLF_QTEDEJAFACT',0);
          //
          if TOBLO.getString('BLO_PIECEPRECEDENTE') <> '' then
          begin
            TOBPDOUV := FindSousDetail (TheAvancLPrec,TOBLO.getString('BLO_PIECEPRECEDENTE'));
            if TOBPDOUV <> nil then
            begin
              TOBLO.SetDouble('BLF_MTCUMULEFACT',TOBPDOUV.GetDouble('BLF_MTCUMULEFACT')+TOBLO.geTDouble('BLF_MTSITUATION'));
              TOBLO.SetDouble('BLF_MTDEJAFACT',TOBPDOUV.GetDouble('BLF_MTCUMULEFACT'));
              //
              TOBLO.SetDouble('BLF_QTECUMULEFACT',TOBPDOUV.GetDouble('BLF_QTECUMULEFACT')+TOBLO.geTDouble('BLF_QTESITUATION'));
              TOBLO.SetDouble('BLF_QTEDEJAFACT',TOBPDOUV.GetDouble('BLF_QTECUMULEFACT'));
              //
              TOBLO.SetDouble('BLF_POURCENTAVANC',Pavanc);
            end else
            begin
              TOBLO.SetDouble('BLF_MTCUMULEFACT',Arrondi(TOBLO.GetDouble('BLF_MTMARCHE')*TOBLO.GetDouble('BLF_POURCENTAVANC')/100,DEV.decimale));
              TOBLO.SetDouble('BLF_QTECUMULEFACT',Arrondi(TOBLO.GetDouble('BLF_QTEMARCHE')*TOBLO.GetDouble('BLF_POURCENTAVANC')/100,V_PGI.okdecQ));
            end;
          end;
        end;
      end;
    end;
  end;

  function TraiteAvancPrec (TOBPiece : TOB) : Boolean;
  var II : Integer;
      TOBL,TOBAvancLPrec,TheSitPre :  TOB;
      SQL : String;
      QQ : TQuery;
  begin
    Result := true;
    TOBAvancLPrec := TOB.Create ('LIGNEFAC',nil,-1);
    TRY
      TheSitPre := FindSitPrec (TOBPiece); // recherche de la facture précédente pour avoir les avancements précédents
      //
      if TheSitPre = nil then Exit;
      //
      Result := false;
      for II := 0 to TOBPiece.Detail.count -1 do
      begin
        TOBL := TOBPiece.detail[II];
        if TOBL.GetString('GL_PIECEPRECEDENTE')<>'' then
        begin
          if not FindLigneSitPrec (TheSitPre,TOBL.GetString('GL_PIECEPRECEDENTE'),TOBAvancLPrec) then  continue;
          if Pos(TOBL.GetString('GL_TYPEARTICLE'),'OUV;ARP')>0 then
          begin
            SQL := 'SELECT LIGNEFAC.*,BLO_TYPEARTICLE,BLO_ARTICLE,BLO_LIBELLE,BLO_NUMLIGNE AS NUMLIG,BLO_N1,BLO_UNIQUEBLO,BLO_PIECEPRECEDENTE,BLO_QTEDUDETAIL,BLO_QTEFACT,BLO_FOURNISSEUR '+
                   'FROM LIGNEFAC '+
                   'LEFT JOIN LIGNEOUV ON '+
                   'BLF_NATUREPIECEG=BLO_NATUREPIECEG AND BLF_SOUCHE=BLO_SOUCHE AND BLF_NUMERO=BLO_NUMERO AND BLF_UNIQUEBLO = BLO_UNIQUEBLO '+
                   'WHERE '+
                   'BLF_NATUREPIECEG="'+TOBAvancLPrec.GetString('BLF_NATUREPIECEG')+'" AND '+
                   'BLF_SOUCHE="'+TOBAvancLPrec.GetString('BLF_SOUCHE')+'" AND '+
                   'BLF_NUMERO="'+TOBAvancLPrec.GetString('BLF_NUMERO')+'" AND '+
                   'BLO_NUMLIGNE="'+TOBAvancLPrec.GetString('NUMLIG')+'" AND '+
                   'BLF_UNIQUEBLO<>0 AND BLO_N2=0 '+
                   'ORDER BY BLO_NUMLIGNE';
            QQ := OpenSQL(SQL,True,-1,'',true); 
            if Not QQ.eof then
            begin
              TOBAvancLPrec.LoadDetailDB('LIGNEFAC','','',QQ,false);
              CalculeAvanc (TOBAvancLPrec);
            end;
            ferme (QQ);
          end;
          AppliqueAvancLignePrec (TOBAvancLPrec,TOBL);
          Result := True;
        end;
      end;
    FINALLY
      TOBAvancLPrec.Free;
    END;
  end;


  procedure EcritLigneFac (TOBPiece : TOB);
  var SQL : String;
      II : Integer;
  begin
    Result := false;
    Sql := 'DELETE FROM LIGNEFAC '+
          'WHERE '+
          'BLF_NATUREPIECEG="'+TOBPiece.GetString ('GP_NATUREPIECEG')+'" AND '+
          'BLF_SOUCHE="'+TOBPiece.GetString('GP_SOUCHE')+'" AND '+
          'BLF_NUMERO="'+TOBPiece.GetString('GP_NUMERO')+'"';
    if ExecuteSQL(SQL)=0 then
    begin
      OneRapport.SauveLigMemo('Document '+TOBPiece.GetString('GP_NUMERO')+ ' -> Erreur / Suppression des anciens avancements ');
      V_PGI.IOError := OeUnknown;
    end;
    TOBPiece.SetAllModifie(true);
    for II := 0 to TOBPiece.detail.count -1 do
    begin
      if Not TOBPiece.detail[II].InsertDBByNivel(false) then
      begin
        OneRapport.SauveLigMemo('Document '+TOBPiece.GetString('GP_NUMERO')+ ' -> Erreur / Ecriture nouveaux avancements ');
        V_PGI.IOError := OeUnknown;
      end;
    end;
  end;

  function LoadSituations (TOBPiece : TOB) : boolean;

    function GetSituations (TOBPiece : TOB) : boolean;
    var QQ : TQuery;
    begin
      result := false;
      QQ := OpenSQL('SELECT * FROM BSITUATIONS WHERE BST_SSAFFAIRE="'+TOBPiece.GetString('GP_AFFAIREDEVIS')+'"',True,-1,'',true);
      if not QQ.Eof then
      begin
        TheSituations.LoadDetailDB('BSITUATIONS','','',QQ,false);
        Result := True;
      end;
      ferme (QQ);
    end;

  begin
    Result := false;
    if TheSituations.Detail.count > 0 then
    begin
      if TOBPiece.GetString('GP_AFFAIREDEVIS') <> TheSituations.detail[0].GetString('BST_SSAFFAIRE') then
      begin
        TheSituations.ClearDetail;
        result := GetSituations (TOBPiece);
      end else Result := true;
    end else
    begin
      result := GetSituations (TOBPiece);
    end;
  end;


begin
  Result := false;
  TRY
    if not LoadSituations (TOBPiece) then
    BEGIN
      OneRapport.SauveLigMemo('Document '+TOBPiece.GetString('GP_NUMERO')+ ' -> Erreur / Récupération des situations précédentes ');
      exit;
    end;
    if not LoadPiece (TOBPiece) then
    begin
      OneRapport.SauveLigMemo('Document '+TOBPiece.GetString('GP_NUMERO')+ ' -> Erreur / Récupération de la situation ');
      exit;
    end;
    if not LoadLesOuvrages(TOBPiece) then
    begin
      OneRapport.SauveLigMemo('Document '+TOBPiece.GetString('GP_NUMERO')+ ' -> Erreur / Récupération des détails d''ouvrages ');
      Exit;
    end;
    if not TraiteAvancPrec (TOBPiece) then
    begin
      OneRapport.SauveLigMemo('Document '+TOBPiece.GetString('GP_NUMERO')+ ' -> Erreur / Récupération de l''avancement précédent ');
      Exit;
    end;
    BEGINTRANS;
    try
      EcritLigneFac (TOBPiece);
      COMMITTRANS;
      Result := True;
    except
      OneRapport.SauveLigMemo('Document '+TOBPiece.GetString('GP_NUMERO')+ ' -> Erreur : écritures des avancements ');
      ROLLBACK;
    end;
  finally
    TOBPiece.clearDetail;
  end;
end;

Initialization
  registerclasses ( [ TOF_BTREAJPIECE_MUL ] ) ; 
	RegisterAglProc('ReajusteAnterieurs', True , 0, AglBTReajusteanterieur);
end.
