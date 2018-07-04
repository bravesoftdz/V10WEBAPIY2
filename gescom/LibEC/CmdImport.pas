unit CmdImport;

interface
uses Classes, UTOF, UTOB, HEnt1, ETransUtil,
{$IFDEF EAGLCLIENT}
      eMul,MaineAGL,
{$ELSE}
      Fiche, HDB, mul, DBGrids, db,dbTables,Fe_Main,
{$ENDIF}
      UxmlUtils, SaisUtil, UtilGrp;

type
    TOF_ImportCmd = class(TOF)
    private
      FMotha : TOB;
      procedure TIntegreCommandes;

    published
      procedure ImportCommand;

    public
      FDeleteOnly : boolean;

    end;

    
function ioErrMsg(ioe : TIOErr; S : string) : String;
procedure IntegreCommandes(SelectedEPieces : TOB; DeleteOnly : boolean; Quiet : boolean = false);


implementation
uses Forms, sysutils, Controls, Dialogs, HCtrls, HMsgBox, HStatus, M3FP,
     Facture, FactGrp, FactCalc, FactComm, FactUtil, FactCpta, FactNomen,
     EntGC, utilPGI, UTofBatchValid, FactTOB, FactTiers, FactPieceContainer;


procedure AGLImportCmd(Parms : Array of Variant; Nb : Integer);
var F : TForm;
    TOTOF : TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFmul) then TOTOF := TFMul(F).LaTOF
                  else exit;
  if (TOTOF is TOF_ImportCmd) then
  begin
    TOF_ImportCmd(TOTOF).FDeleteOnly := Parms[1];
//    Transactions(TOF_ImportCmd(TOTOF).ImportCommand, 1)
    TOF_ImportCmd(TOTOF).ImportCommand;
  end
   else exit;
end;


procedure CreateSoucheIfNeeded(CS : String);
begin
  if ExisteSQL('SELECT SH_SOUCHE FROM SOUCHE WHERE SH_TYPE="GES" AND SH_SOUCHE="'+CS+'"') then exit;
  ExecuteSQL('INSERT INTO SOUCHE (SH_TYPE, SH_SOUCHE, SH_LIBELLE, SH_ABREGE, SH_NUMDEPART, SH_SOCIETE, SH_DATEDEBUT, SH_DATEFIN, SH_FERME, SH_ANALYTIQUE, SH_SIMULATION, SH_NUMDEPARTS, SH_NUMDEPARTP, SH_SOUCHEEXO, SH_RESERVEWEB) '+
                         'VALUES ("GES", "'+CS+'", "Compteur Commandes E-Commerce", "E-Commerce", 1, "'+V_PGI.CodeSociete+'", "1/1/1900", "1/1/1900", "-", "-", "-", 1, 1, "-", "X")');
end;

procedure UpdateNumDepartSouche(CS : String);
var Q : TQuery;
    numd : integer;
begin
  Q := OpenSQL('SELECT MAX(GP_NUMERO) FROM PIECE WHERE GP_SOUCHE="'+CS+'"', true);
  if Q.EOF then numd := 0
           else numd := Q.Fields[0].AsInteger+1;
  Ferme(Q);

  ExecuteSQL('UPDATE SOUCHE SET SH_NUMDEPART='+inttostr(numd)+' WHERE SH_TYPE="GES" AND SH_SOUCHE="'+CS+'"');
end;

procedure RecopieFamilles(TOBArticles, TOBPiece : TOB);
var i, f : integer;
    TOBA : TOB;
begin
   for i := 0 to TOBPiece.Detail.Count-1 do
    if TOBPiece.Detail[i].GetValue('GL_TYPELIGNE')='ART' then
    begin
      TOBA := TOBArticles.FindFirst(['GA_ARTICLE'], [TOBPiece.Detail[i].GetValue('GL_ARTICLE')], false);
      if TOBA = nil then raise Exception.Create('(Recopie familles) Article "'+TOBPiece.Detail[i].GetValue('GL_ARTICLE')+'" non trouvé');
      for f := 1 to 3 do
       TOBPiece.Detail[i].PutValue('GL_FAMILLENIV'+inttostr(f), TOBA.GetValue('GA_FAMILLENIV'+inttostr(f)));
    end;
end;

procedure NomenMoulinette(TOBPiece, TOBArticles, TOBNomenclature : TOB);
var i : integer;
begin
   for i := 0 to TOBPiece.Detail.Count-1 do
    TraiteLesNomenclatures(TOBPiece, TOBArticles, TOBNomenclature, i+1, true);
end;

procedure IntegreCommandes(SelectedEPieces : TOB; DeleteOnly : boolean; Quiet : boolean = false);
var i, j : integer;
    FQ : TQuery;
    MR, Nat, Sou, Nb, Ind : String;
    Suches : Array of String;
    BigMamma, TobTmp: Tob;
    PieceContainer: TPieceContainer;
begin
    if SelectedEPieces.Detail.Count = 0 then begin LogDebugMsg('   Rien à intégrer...', true); exit; end;

    // Grrr à cause des warnings de merde
    if not Assigned(PieceContainer) then
      PieceContainer := TPieceContainer.Create;
    BigMamma := nil;
    PieceContainer.TCArticles := nil;
    PieceContainer.TCBases := nil;
    PieceContainer.TCTiers := nil;
    PieceContainer.TCEches := nil;
    PieceContainer.TCAcomptes := nil;
    PieceContainer.TCCataLogu := nil;
    PieceContainer.TCCpta := nil;
    PieceContainer.TCAnaP := nil;
    PieceContainer.TCAnaS := nil ;
    PieceContainer.TCNomenclature := nil;
    PieceContainer.TCPorcs := nil;

    if not DeleteOnly then
    begin
      SetLength(Suches, 0);
      BigMamma := TOB.Create('Mama_PIECE', nil, -1);

      PieceContainer.TCArticles := TOB.Create('', nil, -1);
      PieceContainer.TCBases := TOB.Create('BASES', nil, -1);
      PieceContainer.TCTiers := TOB.Create('TIERS', nil, -1); PieceContainer.TCTiers.AddChampSup('RIB', False);
      PieceContainer.TCEches := TOB.Create('LES ECHEANCES', nil, -1);
      PieceContainer.CreateTCAcomptes;
      PieceContainer.TCCpta := TOB.Create('', nil, -1);
      PieceContainer.TCAnaP := TOB.Create('', nil, -1);
      PieceContainer.TCAnaS := TOB.Create('', nil, -1);
      //PieceContainer.TCNomenclature := TOB.Create('', nil, -1);
      PieceContainer.TCPorcs := TOB.Create('PIEDPORT', nil, -1);
      PieceContainer.TCCataLogu := TOB.Create('', nil, -1);

      ETableToTable(SelectedEPieces, BigMamma); // BigMamma.detail = les pièces séléctionnées
    end;

    for i := 0 to SelectedEPieces.Detail.Count-1 do
    begin
      Nat := SelectedEPieces.Detail[i].GetValue('EP_NATUREPIECEG');
      Sou := SelectedEPieces.Detail[i].GetValue('EP_SOUCHE');
      Nb := SelectedEPieces.Detail[i].GetValue('EP_NUMERO');
      Ind := SelectedEPieces.Detail[i].GetValue('EP_INDICEG');

      FQ := OpenSQL('SELECT * FROM ELIGNE WHERE EL_NATUREPIECEG="'+Nat+'" '+
                    'AND EL_SOUCHE="'+Sou+'" AND EL_NUMERO='+Nb+' AND EL_INDICEG='+Ind, true);

      SelectedEPieces.Detail[i].LoadDetailDB('ELIGNE', '', '', FQ, false, true);
      Ferme(FQ);

      //Renumérotation lignes
      for j := 0 to SelectedEPieces.Detail[i].Detail.Count-1 do
       SelectedEPieces.Detail[i].Detail[j].PutValue('EL_NUMLIGNE', j+1);

      if not DeleteOnly then
      begin
        PieceContainer.TCPiece := BigMamma.Detail[i];
        ETableToTable(SelectedEPieces.Detail[i], PieceContainer.TCPiece);  // Le détail de chaque piece = lignes de la piece

        for j := 0 to PieceContainer.TCPiece.Detail.Count-1 do
         TOB.Create('', PieceContainer.TCPiece.Detail[j], -1) ; // Ajout 11/01/02

        AddLesSupEntete(PieceContainer.TCPiece);
        PieceAjouteSousDetail(PieceContainer.TCPiece);

        PieceContainer.TCPiece.PutValue('GP_RECALCULER', 'X');
        for j := 0 to PieceContainer.TCPiece.Detail.Count-1 do
         PieceContainer.TCPiece.Detail[j].PutValue('GL_RECALCULER', 'X');

        PieceContainer.DeV.Code := PieceContainer.TCPiece.GetValue('GP_DEVISE');
        GetInfosDevise(PieceContainer.DEV);
        MR := PieceContainer.TCPiece.GetValue('GP_MODEREGLE');
        if RemplirTOBTiers(PieceContainer.TCTiers, PieceContainer.TCPiece.GetValue('GP_TIERS'), Nat, False) <> trtOk then
         begin V_PGI.IoError := oeTiers; break; end;

        try
          UG_AjouteLesArticles(PieceContainer.TCPiece, PieceContainer.TCArticles, PieceContainer.TCCpta, PieceContainer.TCTiers, PieceContainer.TCAnaP, PieceContainer.TCAnaS, PieceContainer.TCCataLogu, false);
          RecopieFamilles(PieceContainer.TCArticles, PieceContainer.TCPiece);

        except
          on e : Exception do begin LogDebugMsg(e.Message, true); V_PGI.IOError := oeStock; break; end;
        end;

        PieceContainer.CleDoc.NaturePiece := PieceContainer.TCPiece.GetValue('GP_NATUREPIECEG');
        PieceContainer.CleDoc.Souche := PieceContainer.TCPiece.GetValue('GP_SOUCHE');
        PieceContainer.CleDoc.NumeroPiece := PieceContainer.TCPiece.GetValue('GP_NUMERO');
        PieceContainer.CleDoc.Indice := PieceContainer.TCPiece.GetValue('GP_INDICEG');
        LoadLesNomen(PieceContainer);
        //NomenMoulinette(PieceContainer.TCPiece, PieceContainer.TCArticles, PieceContainer.TCNomenclature);

        ValideLaPeriode(PieceContainer.TCPiece);
        //Création PieceContainer.TCPorcs
        FQ := OpenSQL('SELECT * FROM PIEDPORT WHERE ' + WherePiece(PieceContainer.CleDoc, ttdPorc, False), True);
        PieceContainer.TCPorcs.LoadDetailDB('PIEDPORT', '', '', FQ, False);
        Ferme(FQ);
        CalculFacture(PieceContainer);
        CalculeSousTotauxPiece(PieceContainer.TCPiece);
        {Visa}
        if (GetInfoParPiece(Nat,'GPP_VISA') = 'X'){NeedVisa}
           and (Abs(PieceContainer.TCPiece.GetValue('GP_TOTALHT')) >= GetInfoParPiece(Nat,'GPP_MONTANTVISA'){MontantVisa})
            then PieceContainer.TCPiece.PutValue('GP_ETATVISA','ATT');

        if ValidationAuto and (not IsModeReglChq(MR)) then
        begin
          // EnregistreAcomptes ici
          TobTmp := TOB.Create('', PieceContainer.TCAcomptes, -1);
          TobTmp.dupliquer(AcompteEnreg(PieceContainer.TCPiece, PieceContainer.TCTiers, Quiet), True, true,true);
          AcomptesVersPiece(PieceContainer);
          if V_PGI.IoError <> oeOK then V_PGI.IoError := oePointage; // oePointage utilisé pour Acompte non enreg.
        end;

        {Echéances}
        PieceContainer.TCEches.ClearDetail;
        GereEcheancesGC(PieceContainer, False);

        PieceContainer.TCPiece.SetAllModifie(True); PieceContainer.TCPiece.SetDateModif(NowH);
        PieceContainer.TCBases.SetAllModifie(True);
        PieceContainer.TCEches.SetAllModifie(True);
        if (PieceContainer.TCAcomptes<>Nil) then PieceContainer.TCAcomptes.SetAllModifie(True);
        PieceContainer.TCTiers.SetAllModifie(True);

        {Enregistrement physique}
        CreateSoucheIfNeeded(Sou);
        if {(Length(Suches)=0) or} (not StrInArray(Sou, Suches)) then
          begin SetLength(Suches, Length(Suches)+1);
                Suches[High(Suches)] := Sou; end;
        if V_PGI.IoError=oeOk then ValideLesLignes(PieceContainer, False, True);
        if (V_PGI.IoError=oeOk) and ValidationAuto and (not IsModeReglChq(MR)) then
          GenereCompta(PieceContainer, PieceContainer.DEV);
        if V_PGI.IoError=oeOk then ValideLesArticles(PieceContainer, PieceContainer.TCPiece) ;
        if V_PGI.IoError=oeOk then ValideleTiers(PieceContainer);
        if V_PGI.IoError=oeOk then PieceContainer.TCBases.InsertOrUpdateDBTable;
        if V_PGI.IoError=oeOk then PieceContainer.TCEches.InsertOrUpdateDBTable;{InsertDBTable(nil);}
        //if V_PGI.IoError=oeOk then ValideLesNomen(PieceContainer.TCNomenclature);
        if (V_PGI.IoError=oeOk) And (PieceContainer.TCAcomptes<>Nil) then PieceContainer.TCAcomptes.InsertOrUpdateDBTable;
        if (PieceContainer.TCAcomptes<>Nil) then PieceContainer.TCAcomptes.Free;

        if MR <> '' then PieceContainer.TCPiece.PutValue('GP_MODEREGLE', MR);

        if V_PGI.IoError=oeOk then PieceContainer.TCPiece.InsertOrUpdateDBTable;
      end{if not deleteonly};

      if V_PGI.IoError=oeOk then
      begin LogDebugMsg('Effacement de pièce : NATUREPIECEG="'+Nat+'" '+' SOUCHE="'+Sou+'" NUMERO='+Nb+' INDICEG='+Ind+' ('+inttostr(SelectedEPieces.Detail[i].Detail.Count)+' lignes)', true);
            for j := 0 to SelectedEPieces.Detail[i].Detail.Count-1 do
              LogDebugMsg('  Ligne '+IntToStr(SelectedEPieces.Detail[i].Detail[j].GetValue('EL_NUMERO'))+' ARTICLE="'+SelectedEPieces.Detail[i].Detail[j].GetValue('EL_ARTICLE')+'"', true);
            if not SelectedEPieces.Detail[i].DeleteDB then LogDebugMsg('Pièce non effacée ! DeleteDB a retourné FAUX', true); end
      else LogDebugMsg('V_PGI Pièce non effacée : NATUREPIECEG="'+Nat+'" '+' SOUCHE="'+Sou+'" NUMERO='+Nb+' INDICEG='+Ind, true);

      if V_PGI.IOError <> oeOK then break;
    end{for (selectedepieces.detail)};

//    if V_PGI.IoError=oeOk then
//     for i := 0 to SelectedEPieces.Detail.Count-1 do SelectedEPieces.Detail[i].DeleteDB(true);

    if not DeleteOnly then
    begin
//      if V_PGI.IoError=oeOk then
//       for i := 0 to BigMamma.Detail.Count-1 do
//        BigMamma.Detail[i].InsertOrUpdateDBTable(true);

      {if V_PGI.IoError=oeOk then} for i := Low(Suches) to High(Suches) do UpdateNumDepartSouche(Suches[i]);

      PieceContainer.TCArticles.Free;
      PieceContainer.TCBases.Free;
      PieceContainer.TCTiers.Free;
      PieceContainer.TCEches.Free;
      PieceContainer.TCCpta.Free;
      PieceContainer.TCAnaP.Free;
      PieceContainer.TCAnas.Free;
      if PieceContainer.TCNomenclature <> nil then PieceContainer.TCNomenclature.Free;
      PieceContainer.TCPorcs.Free;
      PieceContainer.TCCataLogu.Free;

      BigMamma.Free;
      PieceContainer.Free;
    end;
end;

function ioErrMsg(ioe : TIOErr; S : string) : String;
begin
    case ioe of
           oeOK : result := 'Pièce(s) '+S+'ée(s) avec succès';
      oeUnknown : result := 'Erreur inconnue !'#13'Pièce(s) non '+S+'ée(s)';
        oeTiers : result := 'Erreur avec un Tiers !'#13'Pièce(s) non '+S+'ée(s)';
        oeStock : result := 'Erreur avec un Article !'#13'Pièce(s) non '+S+'ée(s)';
     oePointage : result := 'Erreur : Acompte/Règlement non enregistré'; // oePointage utilisé pour Acompte non enreg.
     oeLettrage : result := 'Erreur avec "PassationComptable" !'#13'Pièce(s) non '+S+'ée(s)';

     else result := 'Une erreur s''est produite !'#13'Pièce(s) non '+S+'ée(s)';
    end;
end;

procedure TOF_ImportCmd.TIntegreCommandes;
begin
   IntegreCommandes(FMotha, FDeleteOnly);
end;

procedure TOF_ImportCmd.ImportCommand;
var i : integer;
    FQ : TQuery;
    S, TokKey  : String;
    ioe : TIOErr;
begin
  with TFMul(Ecran) do
  begin
    if FDeleteOnly then S := 'supprim'
                   else S := 'intégr';
    if (FListe.NbSelected = 0) and (not FListe.AllSelected) then
    begin
{$IFNDEF EAGLCLIENT}
    if VAlerte<>Nil then VAlerte.Visible:=FALSE ;
{$ENDIF}
      PGIBox('Veuillez sélectionner les commandes à '+S+'er', Caption);
      exit;
    end;

    if PGIAsk(UpCase(S[1])+Copy(S,2,Length(S))+'er les commandes sélectionnés ?', Caption) = mrNo then Exit;

    FMotha := TOB.Create('Mama_EPIECE', nil, -1);

    Q.DisableControls;
{$IFNDEF EAGLCLIENT}
    if FListe.AllSelected then
    begin
      FQ := PrepareSQL('SELECT * FROM EPIECE', true);
      RecupWhereSQL(Q, FQ);
      FQ.Open;

      FMotha.LoadDetailDB('EPIECE', '', '', FQ, false, true);

      Ferme(FQ);

      FListe.AllSelected := false;
      bSelectAll.Down := false;
    end else
{$ENDIF}
    begin
      InitMove(FListe.NbSelected,'');
      for i := 0 to FListe.NbSelected-1 do
      begin
        FListe.GotoLeBookMark(i);
{$IFDEF EAGLCLIENT}
        Q.TQ.Seek(FListe.Row-1) ;
{$ENDIF}
        TokKey := '"'+Q.{$IFDEF EAGLCLIENT}FindField{$ELSE}FieldByName{$ENDIF}('EP_NATUREPIECEG').AsString+'";'+
                  '"'+Q.{$IFDEF EAGLCLIENT}FindField{$ELSE}FieldByName{$ENDIF}('EP_SOUCHE').AsString+'";'+
                      Q.{$IFDEF EAGLCLIENT}FindField{$ELSE}FieldByName{$ENDIF}('EP_NUMERO').AsString+';'+
                      Q.{$IFDEF EAGLCLIENT}FindField{$ELSE}FieldByName{$ENDIF}('EP_INDICEG').AsString;

        TOB.Create('EPIECE', FMotha, -1).SelectDB(TokKey, nil);

        MoveCur(False);
      end;
      FiniMove;

      FListe.ClearSelected;
    end;
    Q.First;
    Q.EnableControls;

    GetParams;
    //IntegreCommandes(FMotha, FDeleteOnly);
//    if Transactions(TIntegreCommandes, 1) = oeOK then PGIInfo('Commande(s) '+S+'ée(s) avec succès', Caption)
//                                                 else PGIInfo('Erreur, commande(s) non '+S+'ée(s)', Caption);
    ioe := Transactions(TIntegreCommandes, 1);
    PGIInfo(ioErrMsg(ioe, S), Caption);
    //TIntegreCommandes;
    {case ioe of
           oeOK : PGIInfo('Commande(s) '+S+'ée(s) avec succès', Caption);
      oeUnknown : PGIInfo('Erreur inconnue !'#13'Commande(s) non '+S+'ée(s)', Caption);
        oeTiers : PGIInfo('Erreur : Tiers inconnu !'#13'Commande(s) non '+S+'ée(s)', Caption);
        oeStock : PGIInfo('Erreur : Article inconnu !'#13'Commande(s) non '+S+'ée(s)', Caption);
     oeLettrage : PGIInfo('Erreur avec "PassationComptable" !'#13'Commande(s) non '+S+'ée(s)', Caption);

     else PGIInfo('Une erreur s''est produite !'#13'Commande(s) non '+S+'ée(s)', Caption);
    end;}

    FMotha.Free;

    TFMul(Ecran).BChercheClick(Ecran);
//    PGIInfo('Commande(s) '+S+'ée(s) avec succès', Caption);
  end;
end;


initialization
RegisterClasses([TOF_ImportCmd]);
RegisterAGLProc('ImportCmd', true, 1, AGLImportCmd);

end.
