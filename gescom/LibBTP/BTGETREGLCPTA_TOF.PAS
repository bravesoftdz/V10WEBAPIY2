{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 08/12/2015
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTGETREGLCPTA ()
Mots clefs ... : TOF;BTGETREGLCPTA
*****************************************************************}
Unit BTGETREGLCPTA_TOF ;

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
     Paramsoc,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Ent1,
     HTB97,
     vierge,
     UTOF,
     ed_tools,
     UEntCommun,
     UtilsRapport,
     BTPUtil,
     SAISUtil,
     EntGc ;

Type
  TOF_BTGETREGLCPTA = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    OneDateSt : string;
    JusQua : string;
    JournAuxSt : String;
    DateLancement : string;
    TOBReglements,TOBPIECES : TOB;
    RapportRecup :  TGestionRapport;
    function ConstitueDebExerciceCourant : String;
    procedure TraitementRecuperationReglements;
    procedure GetPiecesRegl(TOBR: TOB);
    function ConstitueJournal: string;
    procedure EnregistreReglPiece(TOBReglFAC : TOB);
    procedure ConstitueReglPiece(TOBR : TOB);
    procedure ChargePiece(TOBpiece, TOBAcomptes: TOB; Cledoc: r_cledoc);
    procedure VerifAcomptes(TOBAcomptes: TOB);
    procedure InscritAcomptes(TOBPiece, TOBAcomptes, TOBReglFac: TOB);
    function ExisteRegl(TOBACC: TOB): boolean;
    procedure NettoieAcomptes(TOBAcomptes: TOB);
    procedure BeforeEcritAcomptes(TOBAcomptes,TOBPiece: TOB);
    procedure EcritInfosDoc(TOBAcomptes, TOBPiece: TOB);
    procedure EnregMtSituation(TObacomptes, TOBpiece: TOB);
    procedure EcritInfoAttachements (RefPiece : string;TOBAcomptes : TOB);
  end ;

Implementation

uses DateUtils,factcpta,UtilTOBPiece,UspecifPOC;

procedure TOF_BTGETREGLCPTA.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTGETREGLCPTA.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTGETREGLCPTA.OnUpdate ;
var yy,mm,dd : word;
    Io : TIOErr; 
begin
  Inherited ;
  RapportRecup.InitRapport;
  if ThEdit(getControl('DATEDEPART')).Text <> DateToStr(StrToDateTime(OneDateSt)) then
  begin
    DecodeDate(StrtoDate(ThEdit(getControl('DATEDEPART')).Text),yy,mm,dd);
    DateLancement := DateTimeToStr(EncodeDateTime(yy,mm,dd,0,0,0,0));
  end else
  begin
    DateLancement := OneDateSt;
  end;
  //
  JUSqua := DateTimetoStr(NowH); // prochaine date de mise a jour
  if PgiAsk ('ATTENTION : Vous allez lancer un traitement qui peut être long.#13#10 Confirmez-Vous ?') <> mryes then
  begin
    TFVierge(Ecran).ModalResult := 0;
    exit;
  end;

  Io := Transactions (TraitementRecuperationReglements,0);
  Case io of
    OeOk : begin
            RapportRecup.AfficheRapport;
            SetParamSoc('SO_BTGETREGLLASTDATE',JUSQUA);
            PgiInfo ('Traitement Effectué avec succès');
           end;
    else
      PgiInfo ('Erreur en cours de traitement des Données');
  end;
  TForm(ecran).ModalResult := 0;
end ;

procedure TOF_BTGETREGLCPTA.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTGETREGLCPTA.OnArgument (S : String ) ;
begin
  Inherited ;

  RapportRecup   := TGestionRapport.Create(TForm(ecran));
  RapportRecup.Titre   := 'Récupération des règlements des factures suivantes';
  RapportRecup.Affiche := True;
  RapportRecup.Close   := True;
  RapportRecup.Sauve   := false;
  RapportRecup.Print   := false;
  RapportRecup.InitRapport;
  RapportRecup.Visible := false;
  RapportRecup.PosLeft := Round(ecran.Width / 1.5);
  RapportRecup.PosTop  := Round(ecran.Top);
  RapportRecup.Width := round(screen.width / 2);

  TOBReglements := TOB.Create ('LES REGELEMENST',nil,-1);
  TOBPIECES := TOB.Create ('LES PIECES',nil,-1);
  OneDateSt := GetParamSocSecur('SO_BTGETREGLLASTDATE','');
  JournauxSt := GetParamSocSecur('SO_BTGETJNLCOMPTA','');
  if OneDateSt = '' then
  begin
    OneDateSt := ConstitueDebExerciceCourant;
  end;
  ThEdit(getControl('DATEDEPART')).Text := DateToStr(StrToDateTime(OneDateSt));
  THMultiValComboBox(getControl('JOURNAUX')).Value:= JournauxSt;
end ;

procedure TOF_BTGETREGLCPTA.OnClose ;
begin
  RapportRecup.free;
  TOBReglements.free;
  TOBPIECES.free;
  Inherited ;
end ;

procedure TOF_BTGETREGLCPTA.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTGETREGLCPTA.OnCancel () ;
begin
  Inherited ;
end ;

function TOF_BTGETREGLCPTA.ConstitueDebExerciceCourant: String;
var TheDate : Tdatetime;
    yy,mm,dd : Word;
begin
  TheDate := GetEnCours.Deb;
  DecodeDate (TheDate,yy,mm,dd);
  TheDate := EncodeDateTime (yy,mm,dd,0,0,0,0);
  result := DateTimeToStr(TheDate);
end;


function TOF_BTGETREGLCPTA.ConstitueJournal : string;
var StIn,StST : string;
begin
  StIn := THMultiValComboBox(getControl('JOURNAUX')).Value;
  Result := '';
  repeat
    StST := READTOKENST(StIn);
    if StSt <> '' then
    begin
      if result <> '' then result := Result + ',"'+StSt+'"'
                      else result := '"'+StST+'"';
    end;
  until StSt = '';
end;

procedure TOF_BTGETREGLCPTA.TraitementRecuperationReglements;
var QQ : TQuery;
    Req : String;
    II : integer;
    cledoc : r_cledoc;
begin
  V_PGI.IOError := oeOk;
  //
  TRY
    if VH_GC.BTCODESPECIF = '001' then
    begin
      Req := 'SELECT E_GENERAL, E_AUXILIAIRE, E_LETTRAGE '+
             'FROM ECRITURE '+
             'LEFT OUTER JOIN TIERS ON E_AUXILIAIRE=T_AUXILIAIRE '+
             'WHERE '+
             'E_QUALIFPIECE="N" AND '+
             'E_TRESOLETTRE<>"X" AND '+
             'E_NUMECHE>0 AND '+
             'E_ECHE="X" AND '+
             'E_ETATLETTRAGE<>"R" AND '+
             '((E_DEBIT+E_CREDIT)<>0 OR (E_DEBITDEV+E_CREDITDEV)<>0) AND '+
             'E_ETATLETTRAGE<>"AL" AND '+
             '(E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H") AND '+
             'E_CONFIDENTIEL<="0" AND '+
             'E_DATECOMPTABLE>="'+USDATETIME (StrToDateTime(DateLancement))+'" AND '+
             '(E_GENERAL IN (SELECT G_GENERAL FROM GENERAUXMS WHERE G_NATUREGENE IN ("COC","COF"))) AND '+
             '(E_TIERSPAYEUR="" OR E_TIERSPAYEUR IS NULL) '+
             'GROUP BY E_GENERAL, E_AUXILIAIRE, E_LETTRAGE '+
             'ORDER BY E_GENERAL, E_AUXILIAIRE, E_LETTRAGE';
    end else
    begin
      Req := 'SELECT E_GENERAL, E_AUXILIAIRE, E_LETTRAGE '+
             'FROM ECRITURE '+
             'LEFT OUTER JOIN TIERS ON E_AUXILIAIRE=T_AUXILIAIRE '+
             'WHERE '+
             'E_QUALIFPIECE="N" AND '+
             'E_TRESOLETTRE<>"X" AND '+
             'E_NUMECHE>0 AND '+
             'E_ECHE="X" AND '+
             'E_ETATLETTRAGE<>"R" AND '+
             '((E_DEBIT+E_CREDIT)<>0 OR (E_DEBITDEV+E_CREDITDEV)<>0) AND '+
             'E_ETATLETTRAGE<>"AL" AND '+
             '(E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H") AND '+
             'E_CONFIDENTIEL<="0" AND '+
             'E_DATECOMPTABLE>="'+USDATETIME (StrToDateTime(DateLancement))+'" AND '+
             '(E_GENERAL IN (SELECT G_GENERAL FROM GENERAUXMS WHERE G_NATUREGENE = "COC")) AND '+
             '(E_TIERSPAYEUR="" OR E_TIERSPAYEUR IS NULL) '+
             'GROUP BY E_GENERAL, E_AUXILIAIRE, E_LETTRAGE '+
             'ORDER BY E_GENERAL, E_AUXILIAIRE, E_LETTRAGE';
    end;
    QQ := OpenSql (REq,true,-1,'',true);
    if not QQ.eof then
    begin
      TOBReglements.LoadDetailDB('UN REGLEMENT','','',QQ,false);
    end;
    // ----
    InitMoveProgressForm(nil, 'Analyse des lettrages', 'Veuillez patienter SVP ...', TOBReglements.detail.count, FALSE, TRUE);
    //
    TRY
      if TOBReglements.detail.count > 0 then
      begin
        For II := 0 to TOBReglements.detail.count -1 do
        begin
          MoveCurProgressForm('Analyse des règlements du tiers ' + TOBReglements.detail[II].GetString('E_AUXILIAIRE')+' : Lettrage '+TOBReglements.detail[II].GetString('E_LETTRAGE'));

          GetPiecesRegl(TOBReglements.detail[II]);
          if V_PGI.IOError  <> oeOk then break;
        end;
      end;
    FINALLY
      FiniMoveProgressForm();
    END;
    //
    if V_PGI.ioError = Oeok then
    begin
      InitMoveProgressForm(nil, 'Enregistrement des règlements dans les factures', 'Veuillez patienter SVP ...', TOBPIECES.detail.count, FALSE, TRUE);
      TRY
        for II := 0 to TOBPieces.detail.count - 1 do
        begin
          Cledoc := DecodeRefCPGescom (TOBPieces.detail[II].GetString('REFGESCOM'));
          MoveCurProgressForm('Enregistrement sur ' + rechdom('GCNATUREPIECEG',cledoc.NaturePiece,false)+' N° ' + InttoStr(Cledoc.NumeroPiece));
          EnregistreReglPiece (TOBPieces.detail[II]);
          if V_PGI.IOError  <> oeOk then break;
        end;
      FINALLY
        FiniMoveProgressForm();
      END;
    end;
  FINALLY
    ferme (QQ);
  END;
end;

(*
procedure TOF_BTGETREGLCPTA.TraitementRecuperationReglements;
var QQ : TQuery;
    Req : String;
    II : integer;
begin
  V_PGI.IOError := oeOk;
  TRY
    // récupération des mouvements de règlements clients sur la période définie
    Req := 'SELECT ECRITURE.*, E1.E_NUMTRAITECHQ AS NUMCHEQUE '+
           'FROM ECRITURE '+
           'LEFT JOIN ECRITURE E1 ON '+
           'E1.E_ETABLISSEMENT=ECRITURE.E_ETABLISSEMENT AND '+
           'E1.E_SOCIETE=ECRITURE.E_SOCIETE AND '+
           'E1.E_EXERCICE = ECRITURE.E_EXERCICE AND '+
           'E1.E_GENERAL=ECRITURE.E_CONTREPARTIEGEN AND '+
           'E1.E_JOURNAL=ECRITURE.E_JOURNAL AND '+
           'E1.E_QUALIFPIECE=ECRITURE.E_QUALIFPIECE AND '+
           'E1.E_NATUREPIECE=ECRITURE.E_NATUREPIECE AND '+
           'E1.E_NUMEROPIECE=ECRITURE.E_NUMEROPIECE '+
           'WHERE '+
           'ECRITURE.E_LETTRAGE <> "" AND '+
           'ECRITURE.E_DATEMODIF >="'+ USDATETIME (StrToDateTime(DateLancement)) + '" AND '+
           'ECRITURE.E_DATEMODIF <="'+ USDATETIME (StrToDateTime(JusQua)) + '" AND '+
           'ECRITURE.E_NATUREPIECE IN ("RC","OC","OD") AND ECRITURE.E_AUXILIAIRE <> "" AND '+
           'ECRITURE.E_JOURNAL IN ('+ConstitueJournal+')';
    QQ := OpenSql (REq,true,-1,'',true);
    if not QQ.eof then
    begin
      TOBReglements.LoadDetailDB('UN REGLEMENT','','',QQ,false);
    end;
  FINALLY
    ferme (QQ);
  END;
  if TOBReglements.detail.count > 0 then
  begin
    For II := 0 to TOBReglements.detail.count -1 do
    begin
      GetPiecesRegl(TOBReglements.detail[II]);
      if V_PGI.IOError  <> oeOk then break;
    end;
  end;
end;
*)

procedure TOF_BTGETREGLCPTA.GetPiecesRegl (TOBR : TOB);
var Req : String;
    QQ : TQuery;
begin
    Req := 'SELECT * '+
           'FROM ECRITURE '+
           'WHERE '+
           'E_AUXILIAIRE = "'+TOBR.GetString('E_AUXILIAIRE')+'" AND '+
           'E_GENERAL = "'+TOBR.GetString('E_GENERAL')+'" AND '+
           'E_LETTRAGE = "'+TOBR.GetString('E_LETTRAGE')+'" AND '+
//           'E_REFGESCOM <> "" AND NOT E_NATUREPIECE IN ("RC","OC") AND '+
           'E_QUALIFPIECE="N" AND '+
           'E_TRESOLETTRE<>"X" AND '+
           'E_NUMECHE>0 AND '+
           'E_ECHE="X" AND '+
           'E_ETATLETTRAGE<>"RI" AND '+
           '((E_DEBIT+E_CREDIT)<>0 OR (E_DEBITDEV+E_CREDITDEV)<>0) AND '+
           'E_ETATLETTRAGE<>"AL" AND '+
           '(E_ECRANOUVEAU="N" OR '+
           'E_ECRANOUVEAU="H") AND '+
           'E_CONFIDENTIEL<="0" AND '+
           'E_DATECOMPTABLE>="'+USDATETIME (VH^.Precedent.Deb)+'" '+
           'ORDER BY E_AUXILIAIRE, E_GENERAL,E_DATECOMPTABLE ,E_NUMEROPIECE,E_NUMLIGNE,E_NUMECHE';

    QQ := OpenSql (REq,true,-1,'',true);
    TRY
      if not QQ.eof then
      begin
        TOBR.LoadDetailDB('UNE REGLEMENT FAC','','',QQ,false);
        ConstitueReglPiece (TOBR);
        end;
    FINALLY
    ferme(QQ);
    END;
end;

procedure  TOF_BTGETREGLCPTA.ConstitueReglPiece (TOBR : TOB);
var TOBPP,TOBINT,TOBDRD,TOBRP : TOB;
    RefPiece : string;
    I,J : integer;
begin
  I := 0;
  TOBINT := TOB.Create ('LES FACTURES REGLEES',nil,-1);
  TRY
  repeat
      TOBDRD := TOBR.detail[I];
      if (TOBDRD.getString('E_NATUREPIECE')='FC') and (TOBDRD.GEtDouble('E_DEBITDEV')<>0) then
      begin
        refPiece := TOBDRD.GetString('E_REFGESCOM');
        if refpiece <> '' then TOBDRD.ChangeParent(TOBINT,-1)
                          else TOBDRD.free;
      end else if (TOBDRD.getString('E_NATUREPIECE')='FF') and (TOBDRD.GEtDouble('E_CREDITDEV')<>0) then
    begin
        refPiece := TOBDRD.GetString('E_REFGESCOM');
        if refpiece <> '' then TOBDRD.ChangeParent(TOBINT,-1)
                          else TOBDRD.free;
      end else inc(I);
    until I >=TOBR.detail.count;
    //
    if TOBINT.detail.count > 0 then
    begin
      //
      for I := 0 to TOBR.detail.count -1 do
      begin
        TOBDRD := TOBR.detail[I];
        for J := 0 to TOBINT.detail.count -1 do
        begin
          TOBPP := TOB.Create ('ECRITURE',TOBINT.detail[J],-1);
          TOBPP.Dupliquer(TOBDRD,false,true);
    end;
      end;
      //
      I :=0;
      repeat
        TOBRP := TOBPIECES.FindFirst(['REFGESCOM'],[TOBINT.detail[I].getString('E_REFGESCOM')],false);
        if TOBRP= nil then
        begin
          TOBRP := TOB.Create('UNE PIECE',TOBPIECES,-1);
          TOBRP.AddChampSupValeur ('REFGESCOM',TOBINT.detail[0].getString('E_REFGESCOM'));
        end;
        TOBINT.detail[0].ChangeParent(TOBRP,-1);
      until I >= TOBINT.detail.count;
    end;
  FINALLY
    TOBR.ClearDetail;
    TOBINT.free;
  END;
end;

procedure TOF_BTGETREGLCPTA.ChargePiece (TOBpiece,TOBAcomptes: TOB; Cledoc : r_cledoc);
var Sql : string;
    QQ : Tquery;
begin
  Sql := 'SELECT * FROM PIECE WHERE '+WherePiece(Cledoc,ttdPiece,true);
  QQ := OpenSql(SQl,true,1,'',true);
  if not QQ.eof then
  begin
    TOBPiece.SelectDB('',QQ);
  end;
  ferme (QQ);
  Sql := 'SELECT ACOMPTES.*,"X" AS EXISTE,"-" AS ASUPR FROM ACOMPTES WHERE '+
          WherePiece(Cledoc,ttdAcompte,true)+
          ' ORDER BY GAC_JALECR,GAC_NUMECR,GAC_MODEPAIE,GAC_DATEECR';
  QQ := OpenSql(SQl,true,1,'',true);
  if not QQ.eof then
  begin
    TOBAcomptes.LoadDetailDB('ACOMPTES','','',QQ,false);
  end;
  ferme (QQ);
end;

function TOF_BTGETREGLCPTA.ExisteRegl (TOBACC : TOB) : boolean;
var SQl : String;
begin
  SQl := 'SELECT 1 FROM ECRITURE '+
         'WHERE '+
         'E_AUXILIAIRE="'+TOBACC.GetString('GAC_AUXILIAIRE')+'" AND '+
         'E_JOURNAL="'+TOBACC.GetString('GAC_JALECR')+'" AND '+
         'E_NUMEROPIECE='+ TOBACC.GetString('GAC_NUMECR')+ ' AND '+
         'E_MODEPAIE="'+TOBACC.GetString('GAC_MODEPAIE')+'" AND '+
         'E_DATECOMPTABLE="'+USDATETIME(StrToDate(DateToStr(TOBACC.GetDateTime('GAC_DATEECR'))))+'"';
  result := ExisteSQL(SQl);
end;

procedure TOF_BTGETREGLCPTA.VerifAcomptes(TOBAcomptes : TOB);
var II : integer;
begin
  if TOBAcomptes.detail.count = 0 then exit;
  for II := 0 to TOBAcomptes.detail.count -1 do
  begin
    if not ExisteRegl(TOBacomptes.detail[II]) then TOBAcomptes.detail[II].SetString('ASUPR','X');
  end;
end;


procedure TOF_BTGETREGLCPTA.EnregistreReglPiece  (TOBReglFAC : TOB);
var TOBPiece,TOBAcomptes : TOB;
    RefPiece,RefPieceATT : string;
    Cledoc : R_Cledoc;
    MsgDocument : string;
begin
  TOBPiece := TOB.Create('PIECE',nil,-1);
  TOBACOMPTES := TOB.Create ('LES ACOMPTES',nil,-1);
  TRY
    RefPiece := TOBReglFAC.GetString('REFGESCOM');
    Cledoc := DecodeRefCPGescom (RefPiece);
    ChargePiece (TOBpiece,TOBAcomptes,Cledoc);
    if V_PGI.IOError <> OEOk then Exit;
    if TOBPiece.getString('GP_ATTACHEMENT')<> '' then RefPieceATT := TOBPiece.getString('GP_ATTACHEMENT'); // on commence par recup le lien avec la facture d'acompte...  
    VerifAcomptes(TOBAcomptes);
    if V_PGI.IOError <> OEOk then Exit;
    InscritAcomptes (TOBPiece,TOBAcomptes,TOBReglFac);
    if V_PGI.IOError <> OEOk then Exit;
    NettoieAcomptes (TOBAcomptes);
    if V_PGI.IOError <> OEOk then Exit;
    BeforeEcritAcomptes (TOBAcomptes,TOBPiece);
    if V_PGI.IOError <> OEOk then Exit;
    EcritInfosDoc (TOBAcomptes,TOBPiece);
    if V_PGI.IOError <> OEOk then Exit;
    if refPieceATT <> '' then EcritInfoAttachements (RefPieceATT,TOBAcomptes);
    if (VH_GC.BTCODESPECIF = '001') and (TOBPiece.GetString('GP_NATUREPIECEG')='FF') then
    begin
      SetFactureReglePOC (TOBpiece);
    end;
    MsgDocument := 'Reglements pour la '+
                   rechdom('GCNATUREPIECEG',TOBpiece.getValue('GP_NATUREPIECEG'),false)+' N° '+
                   InttoStr(TOBPiece.getInteger('GP_NUMERO'))+' '+
                   'd''un montant de '+
                   StrF00(TOBPiece.getDouble('GP_ACOMPTEDEV'),V_PGI.okdecV); 
    RapportRecup.SauveLigMemo(MsgDocument);
  FINALLY
    TOBPiece.free;
    TOBACOMPTES.free;
  END;
end;

procedure TOF_BTGETREGLCPTA.EnregMtSituation (TObacomptes,TOBpiece : TOB) ;
var Req : string;
    MontantRegl,MontantAcompte : double;
    TOBSIt : TOB;
    QQ : TQuery;
begin
  req := 'SELECT * FROM BSITUATIONS WHERE BST_NATUREPIECE="'+TOBPIece.GetValue('GP_NATUREPIECEG')+'" AND ';
  req := Req + 'BST_SOUCHE="'+TOBPIece.GetVAlue('GP_SOUCHE')+'" AND BST_NUMEROFAC="'+inttoStr(TOBPiece.GetValue('GP_NUMERO'))+'"';
  QQ := OpenSql(Req,true,-1,'',true);
  if not QQ.eof then
  begin
    GetMontantsAcomptes (TOBAcomptes,MontantAcompte,MontantRegl);
    TOBSit := TOB.create ('BSITUATIONS',nil,-1);
    TRY
      TOBSIt.SelectDB ('',QQ);
      TOBSIT.PutValue('BST_MONTANTREGL',MontantRegl);
      TOBSIT.PutValue('BST_MONTANTACOMPTE',MontantAcompte);
      TOBSit.SetAllModifie (true);
      if not TOBSIT.UpdateDB (false) then V_PGI.ioerror := oeUnknown;
    FINALLY
      TOBSIT.free;
    END;
  end;
  ferme (QQ);
end;

procedure TOF_BTGETREGLCPTA.EcritInfosDoc (TOBAcomptes,TOBPiece : TOB);
var II : Integer;
begin
  For II := 0 to TOBAcomptes.detail.count -1 do
  begin
    if TOBAcomptes.detail[II].GetString('EXISTE')='X' then
    begin
      if not TOBAcomptes.detail[II].UpdateDB then
      begin
        V_PGI.ioerror := oeUnknown;
      end;
    end else
    begin
      TOBAcomptes.detail[II].SetAllModifie(true);
      if not TOBAcomptes.detail[II].InsertDB(nil) Then
      begin
        V_PGI.ioerror := oeUnknown;
      end;
    end;
  end;
  //
  EnregMtSituation (TObacomptes,TOBpiece);
  if V_PGI.ioerror = OeOk then if not TOBPiece.UpdateDB then V_PGI.ioerror := oeUnknown;
  
end;


procedure TOF_BTGETREGLCPTA.InscritAcomptes (TOBPiece,TOBAcomptes,TOBReglFac : TOB);
var II : integer;
    TOBRF,TOBACC : TOB;
begin
  For II := 0 to TOBReglFac.detail.count -1  do
  begin
    TOBRF := TOBReglFac.detail[II];
    TOBACC := TOBAcomptes.findFirst(['GAC_JALECR','GAC_NUMECR','GAC_MODEPAIE','GAC_DATEECR'],
                                  [TOBRF.GetString('E_JOURNAL'),
                                   TOBRF.GetInteger('E_NUMEROPIECE'),
                                   TOBRF.GetString('E_MODEPAIE'),
                                   TOBRF.GetDateTime ('E_DATECOMPTABLE')],true);
    if TOBACC = nil then
    begin
      //n'existe pas encore..---> c'est un petit nouveauuuu
      TOBACC := TOB.Create ('ACOMPTES',TOBAComptes,-1);
      TOBACC.AddChampSupValeur('EXISTE','-');
      TOBACC.AddChampSupValeur('ASUPR','-');
      TOBACC.SetString('GAC_NATUREPIECEG',TOBpiece.GetString('GP_NATUREPIECEG'));
      TOBACC.SetString('GAC_SOUCHE',TOBpiece.GetString('GP_SOUCHE'));
      TOBACC.SetInteger('GAC_NUMERO',TOBpiece.GetInteger('GP_NUMERO'));
      TOBACC.SetInteger('GAC_INDICEG',TOBpiece.GetInteger('GP_INDICEG'));
      TOBACC.SetString('GAC_JALECR',TOBRF.GetString('E_JOURNAL'));
      TOBACC.SetInteger('GAC_NUMECR',TOBRF.GetInteger('E_NUMEROPIECE'));
      TOBACC.SetDouble('GAC_MONTANT',TOBRF.GetDouble('E_COUVERTURE'));
      TOBACC.SetDouble('GAC_MONTANTDEV',TOBRF.GetDouble('E_COUVERTUREDEV'));
      TOBACC.SetString('GAC_MODEPAIE',TOBRF.GetString('E_MODEPAIE'));
      if TOBRF.GetString('E_NATUREPIECE')='RC' then TOBACC.SetString('GAC_ISREGLEMENT','X') else TOBACC.SetString('GAC_ISREGLEMENT','-'); 
      TOBACC.SetString('GAC_NUMCHEQUE',TOBRF.getstring('NUMCHEQUE'));
      TOBACC.SetString('GAC_LIBELLE',TOBRF.getstring('E_LIBELLE'));
      TOBACC.SetString('GAC_QUALIFPIECE',TOBRF.getstring('E_QUALIFPIECE'));
      TOBACC.SetString('GAC_AUXILIAIRE',TOBRF.getstring('E_AUXILIAIRE'));
      TOBACC.SetString('GAC_CPTADIFF','-');
      TOBACC.PutValue('GAC_DATEECR',TOBRF.GetValue('E_DATECOMPTABLE'));
      TOBACC.PutValue('GAC_DATEECHEANCE',TOBRF.GetValue('E_DATEECHEANCE'));
    end;
  end;
end;

procedure TOF_BTGETREGLCPTA.NettoieAcomptes (TOBAcomptes : TOB);
var II : integer;
begin
  II := 0;
  repeat
    if TOBAcomptes.detail[II].getString('ASUPR')='X' then
    begin
      TOBAcomptes.detail[II].DeleteDB; 
      TOBAcomptes.detail[II].free;
    end else inc(II);
  until II >= TOBAcomptes.detail.count;
end;

procedure TOF_BTGETREGLCPTA.BeforeEcritAcomptes (TOBAcomptes,TOBPiece : TOB);
var II : integer;
    NumOrdre : integer;
    LJALECR : string;
    LNUMECR : integer;
    JALECR : string;
    NUMECR : integer;
    TOBA : TOB;
    Acompte ,AcompteDev : double;
begin
  TOBAcomptes.detail.sort ('GAC_JALECR,GAC_NUMECR');
  NumOrdre := 0;
  Acompte := 0;
  AcompteDev := 0;
  LJALECR := '';
  LNUMECR := 0;
  for II := 0 to TOBAcomptes.detail.count -1 do
  begin
    TOBA := TOBAcomptes.detail[II];

    JALECR := TOBA.getString('GAC_JALECR');
    NUMECR := TOBA.GetInteger('GAC_NUMECR');

    if (JALECR<>LJALECR) or
       (NUMECR<>LNUMECR) then
    begin
      LJALECR := JALECR;
      LNUMECR := NUMECR;
      NumOrdre := 0;
    end else Inc(Numordre);
    TOBA.setInteger('GAC_NUMORDRE',NumOrdre);
    Acompte := Acompte + TOBA.getDouble ('GAC_MONTANT');
    AcompteDEV := AcompteDEV + TOBA.getDouble ('GAC_MONTANTDEV');
  end;
  TOBPiece.SetDouble('GP_ACOMPTE',Acompte);
  TOBPiece.SetDouble('GP_ACOMPTEDEV',AcompteDEV);
end;

procedure TOF_BTGETREGLCPTA.EcritInfoAttachements(RefPiece: string; TOBAcomptes: TOB);
var TOBACC,TOBAC : TOB;
    I : integer;
    Zcledoc : r_cledoc;
    MtAcc,MtAccDev : double;
begin
  if refPiece = '' then exit;
  DecodeRefPiece(refPiece,ZCledoc);
  TRY
    ExecuteSql ('DELETE FROM ACOMPTES WHERE '+WherePiece (Zcledoc,ttdAcompte,false));
  EXCEPT
    on E: Exception do
    begin
      PgiError('Erreur SQL : ' + E.Message, 'Ecriture acomptes');
      V_PGI.ioerror := oeUnknown;
      exit;
    end;
  END;
  TOBACC := TOB.create ('LES ACOMPTES',nil,-1);
  TOBACC.Dupliquer( TOBAComptes,true,true);
  MtAcc := 0; MtAccDev := 0;
  for I:= 0 to TOBACC.detail.count -1 do
  begin
    TOBAC := TOBACC.detail[I];
    TOBAC.SetString('GAC_NATUREPIECEG',ZCledoc.NaturePiece);
    TOBAC.SetString('GAC_SOUCHE',ZCledoc.Souche);
    TOBAC.SetInteger('GAC_NUMERO',ZCledoc.NumeroPiece);
    TOBAC.SetInteger('GAC_INDICEG',ZCledoc.Indice);
    TOBAC.Setboolean('GAC_ISREGLEMENT',true);
    MtAcc := MtAcc + TOBAC.getDouble('GAC_MONTANT');
    MtAccDEV := MtAccDev + TOBAC.getDouble('GAC_MONTANTDEV');
    TOBAC.SetAllModifie(true);
    TRY
      TOBAC.InsertDB(nil);
    EXCEPT
      on E: Exception do
      begin
        PgiError('Erreur SQL : ' + E.Message, 'Ecriture acomptes');
        V_PGI.ioerror := oeUnknown;
        exit;
      end;
    END;
  end;
  TRY
    ExecuteSql ('UPDATE PIECE SET '+
                'GP_ACOMPTE='+STRfPoint(MtAcc)+', '+
                'GP_ACOMPTEDEV='+STRfPoint(MtAccDEV)+' '+
                'WHERE '+WherePiece(Zcledoc,ttdpiece,false));
  EXCEPT
    on E: Exception do
    begin
      PgiError('Erreur SQL : ' + E.Message, 'Ecriture Piece');
      V_PGI.ioerror := oeUnknown;
      exit;
    end;
  END;

end;

Initialization
  registerclasses ( [ TOF_BTGETREGLCPTA ] ) ;
end.
