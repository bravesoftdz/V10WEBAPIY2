{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 26/06/2003
Modifié le ... : 13/06/2006
Description .. : Source TOF de la FICHE : CPREIMPUT ()
Suite ........ : 
Suite ........ : JP 01/08/05 : FQ 15124 : on interdit les réimputations sur
Suite ........ : les lignes TTC venant de la GESCOM
Suite ........ : 
Suite ........ : JP 22/08/05 : FQ 15931 : Gestion de E_TRESOSYNCHRO
Suite ........ : 
Suite ........ : JP 13/06/06 : FQ 18329 : Gestion de E_TRESOSYNCHRO 
Suite ........ : sur les contreparties des écritures modifiées
Mots clefs ... : TOF;CPREIMPUT
*****************************************************************}
unit CPREIMPUT_TOF;

interface

uses StdCtrls,
  Controls,
  Classes,
  Utob,
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  eMul, // TFMul
  {$ELSE}
  db,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  FE_Main,
  Hdb, // THDBGrid
  Mul, // TFMul
  {$ENDIF}
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  UtilSoc, // MarquerPublifi
  Saisie, // TrouveEtLanceSaisie
  forms,
  sysutils,
  ComCtrls,
  HCtrls, // THGrid
  HEnt1,
  HMsgBox,
  utilPGI, // _Blocage
  LettUtil, // RempliSelectEuro
  Ent1, // MaxAxe, VH
  SAISUTIL, // TOBM, RMVT, GereSelectionsGrid
  DelVisue, // VisuPiecesGenere
  HTB97, // TToolBarButton97
  Ed_Tools, // VideListe
  AGLInit, // TheTob
  HStatus, // InitMove, MoveCur, FiniMove
  SaisComm, // RemplirOFromM
  LetBatch, // VentileGenerale
  ULibEcriture, //CAffectCompteContrePartie  FQ 14849
  ULibAnalytique, //CChargeAna FQ 14849
  ParamSoc,		// GetParamSocSecur YMO
  UTOF,
  SoldeCpt,
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  dialogs;

//procedure Reimputation;
procedure CC_LanceFicheReimp;

type
  TOF_CPREIMPUT = class(TOF)
    HM: THMsgBox;
    Fiche: TFMul;
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnDisplay; override;
    procedure OnClose; override;
    procedure OnCancel; override;
    procedure InitMsgBox;

    procedure E_EXERCICEChange(Sender: TObject);
    procedure BListePIECESClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure BZoomClick(Sender: TObject);
    procedure BChercheClick(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject);
    procedure AuxiElipsisClick(Sender : TObject);

  private
    gbDeviseNotPivot, gbCONTRE_NEGATIF, gbDirect: Boolean;
    GeneDate, NowFutur: TDateTime;
    ListeSel, ListeAna, TPIECE: TList;
    MemeVentil: boolean;
    EncON: boolean;
    Gen2Vent, IsVenti, MsgSection: array[1..MaxAxe] of boolean;
    gszCONTRE_GENERAL, gszCONTRE_AUXILIAIRE, gszCONTRE_JOURNAL, gszCONTRE_ETABLISSEMENT: string;

    function CoherenceOK: boolean;
    procedure ChargeSelection;
    function JePeuxValider: boolean;
    procedure TripoteO(O: TOBM; M: RMVT; EcrG: boolean);
    procedure TraiteLesAna(TAna: TOB; OEcr: TOBM; NewNumL: integer; OldM, NewM: RMVT; Ventilable: Boolean = True);
    procedure ReimputeAna(TAna: TOB);
    procedure ContrePasseAna(TAna: TOB);
    procedure Reimpute(TEcr: TOB);
    procedure ContrePasse(TEcr: TOB);
    procedure GenereLesPieces;
    procedure ReimputeTobAna(Ana: Tob; OEcr: TOBM; Axe: Integer);
    function IsVentilable(General: string): Boolean;
    procedure SetContrePartie(Ecriture: TOBM); //FQ 14849 SG6 04/11/2004
    {JP 01/08/05 : FQ 15124 : On regarde s'il y a dans la sélection une ligne TTC de Gescom}
    function  PresenceTTCGescom : Boolean;
    {JP 22/08/05 : FQ 15931 : Suppression des écritures en Treso et maj de E_TRESOSYNCHRO}
    function MajTreso(T : TOB; Detruire : Boolean) : string;
    {CA 12/07/2006 : FQ 18555 : Correction SelectALL }
    procedure SlctAllClick(Sender : TObject);
    {JP 26/06/07 : FQ TRESO 10491 : on empêche la suppression des écritures issues de la Trésorerie}
    function  PresenceMvtTreso : Boolean;
    procedure MajdesSoldesImp;
  end;


implementation

uses
  {$IFDEF MODENT1}
  CPVersion,
  CPProcMetier,
  ULibExercice,
  CPProcGen,
  {$ENDIF MODENT1}
  CPREIMPUTPARAM_TOF, ULibTrSynchro {DetruitEcritureTresoTOB}, Constantes {ets_Nouveau},
  {$IFDEF eAGLCLIENT}
  MenuOLX
  {$ELSE}
  MenuOLG
  {$ENDIF eAGLCLIENT}
  , UTofMulParamGen; {13/04/07 YMO F5 sur Auxiliaire }

procedure CC_LanceFicheReimp;
begin
  if PasCreerDate(V_PGI.DateEntree) then Exit;
  if _Blocage(['nrBatch', 'nrCloture'], True, 'nrBatch') then Exit;
  AGLLanceFiche('CP', 'CPREIMPUT', '', '', '');
end;

procedure TOF_CPREIMPUT.OnNew;
begin
  inherited;
end;

procedure TOF_CPREIMPUT.OnDelete;
begin
  inherited;
end;

procedure TOF_CPREIMPUT.OnUpdate;
begin
  inherited;
end;

procedure TOF_CPREIMPUT.OnLoad;
var
  s : string;
  Q : TQuery;
begin
  inherited;
  {JP 26/09/05 : FQ 16651 : pour éviter d'affecter un compte bancaire sur un journal de banque
                 dont il ne serait pas le compte de contrepartie, on n'affiche pas les écritures
                 dont le général est le compte de contrepartie}
  s := GetControlText('E_JOURNAL');

  {Valeur par défaut du XX_WHERE2}
  SetControlText('XX_WHERE2', 'E_REFPOINTAGE = ""');

  {Recherche du compte de contrepartie du journal ...}
  Q := OpenSQL('SELECT J_CONTREPARTIE FROM JOURNAL WHERE J_JOURNAL = "' + s + '"', True);
  try
    if not Q.EOF then begin
      s := Q.FindField('J_CONTREPARTIE').AsString;
      {S'il y en a un, on l'exclut du Mul}
      if s <> '' then
        SetControlText('XX_WHERE2', 'E_REFPOINTAGE = "" AND  E_GENERAL <> "' + s + '"');
    end;
  finally
    Ferme(Q);
  end;
end;

procedure TOF_CPREIMPUT.OnArgument(S: string);
var
  E_JOURNAL: THValComboBox;
  year_clotperio,month_clotperio,day_clotperio:word;
  year_aprendre,month_aprendre:word;
begin
  inherited;
  Ecran.HelpContext := 7676000;
  Decodedate(VH^.DateCloturePer,year_clotperio,month_clotperio,day_clotperio);
  if month_clotperio=12 then
  begin
    month_aprendre:=1;
    year_aprendre:=year_clotperio+1;
  end
  else
  begin
    month_aprendre:=month_clotperio+1;
    year_aprendre:=year_clotperio;
  end;

  {JP 22/08/05 : On ne peut modifier les écritures de Tréso en comptabilité}
  SetControlText('XX_WHERE3','(E_QUALIFORIGINE <> "' + QUALIFTRESO + '" OR E_QUALIFORIGINE IS NULL OR E_QUALIFORIGINE = "") AND E_NATUREPIECE<>"ECC" AND E_DATECOMPTABLE>="'+usdatetime(encodedate(year_aprendre,month_aprendre,1))+'"');

  // Création des contrôles
  HM := THMsgBox.Create(FMenuG);
  InitMsgBox;

  // Evénements des contrôles
  E_JOURNAL := THValComboBox(GetControl('E_JOURNAL', True));
  THValComboBox(GetControl('E_EXERCICE', True)).OnChange := E_EXERCICEChange;
  // Ajout me fiche 18990
  THValComboBox(GetControl('E_EXERCICE', True)).plus := 'AND EX_ETATCPTA="OUV"';
  TToolBarButton97(GetControl('BLISTEPIECES', True)).OnClick := BListePIECESClick;
  TToolBarButton97(GetControl('BZOOM', True)).OnClick := BZoomClick;
  TToolBarButton97(GetControl('BCHERCHE', True)).OnClick := BChercheClick;
  TToolBarButton97(GetControl('BOUVRIR', True)).OnClick := BOuvrirClick;
  TToolbarButton97(GetControl('BSELECTALL')).OnClick := SlctAllClick;
  Fiche := TFMul(Ecran);

  {$IFDEF EAGLCLIENT}
  THGrid(GetControl('FLISTE')).OnDblClick := FListeDblClick;
  {$ELSE}
  THDBGrid(GetControl('FLISTE')).OnDblClick := FListeDblClick;
  {$ENDIF}

  //////////////////////////////////////////////////////////////////
  // FormCreate
  ListeSel := TList.Create;
  ListeAna := TList.Create;
  TPIECE := TList.Create;
  EncON := False;
  gbDeviseNotPivot := False;

  //////////////////////////////////////////////////////////////////
  // FormShow
  SetControlText('E_DEVISE', V_PGI.DevisePivot);
  //  SetControlText('CModeS', V_PGI.DevisePivot);
  if GetCPExoRef.Code <> '' then
  begin
    SetControlText('E_EXERCICE', GetCPExoRef.Code);
    SetControlText('E_DATECOMPTABLE', DateToStr(GetCPExoRef.Deb));
    SetControlText('E_DATECOMPTABLE_', DateToStr(GetCPExoRef.Fin));
  end
  else
  begin
    SetControlText('E_EXERCICE', GetEntree.Code);
    SetControlText('E_DATECOMPTABLE', DateToStr(V_PGI.DateEntree));
    SetControlText('E_DATECOMPTABLE_', DateToStr(V_PGI.DateEntree));
  end;
  SetControlText('E_DATEECHEANCE', StDate1900);
  SetControlText('E_DATEECHEANCE_', StDate2099);

  if (E_JOURNAL.Items.Count > 0) then
  begin
    if E_JOURNAL.Vide then E_JOURNAL.ItemIndex := 1
    else E_JOURNAL.ItemIndex := 0;
  end;
  // 51320 : Specif TESSI TVA : possibilité de réimputation sur pièce GESCOM + onglet avancé (CPREIMPUT_TOF.PAS)
  if not EstSpecif('51320') then
    TTabSheet(GetControl('PAVANCE', True)).TabVisible := False;
  Fiche.bSelectAll.Visible := True; // 14686

  PositionneEtabUser(GetControl('E_ETABLISSEMENT'), False); // 15090

  if GetParamSocSecur('SO_CPMULTIERS', false) then
    THEdit(GetControl('E_AUXILIAIRE', true)).OnElipsisClick:=AuxiElipsisClick;

end;

procedure TOF_CPREIMPUT.OnClose;
begin
  inherited;
  HM.Free;
  _Bloqueur('nrBatch', False);

  //////////////////////////////////////////////////////////////////
  // FormClose
  VideListe(ListeSel);
  ListeSel.Free;
  VideListe(ListeAna);
  ListeAna.Free;
  VideListe(TPIECE);
  TPIECE.Free;

end;

procedure TOF_CPREIMPUT.OnDisplay();
begin
  inherited;
end;

procedure TOF_CPREIMPUT.OnCancel();
begin
  inherited;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Stephane Guillon
Créé le ...... : 04/11/2004
Modifié le ... :   /  /
Description .. : Pour un numéro de pièces, set les bonnes contreparties aux
Suite ........ : et gen ainsi que pour les eventuelles ventilations
Suite ........ : analytiques
Mots clefs ... : FQ 14849   SG6 04/11/2004
*****************************************************************}
procedure TOF_CPREIMPUT.SetContrePartie(Ecriture: TOBM);
var
  tobecriture: TOB;
  i, j, z: integer;
begin
  //Chargement écritures pour le numéro de pièce
  tobecriture := TOB.Create('ECRITURES', nil, -1);
  tobecriture.LoadDetailDBFromSQL('ECRITURE', 'SELECT * FROM ECRITURE WHERE ' + WhereEcritureTOB(tsGene, Ecriture, false));


  //Modification des contrepartie AUX ET GEN (pour les ecritures)
  CAffectCompteContrePartie(tobecriture);

  //Parcours de la tobecriture > si modifier > modification des analytiques
  for i := 0 to tobecriture.Detail.Count - 1 do
  begin
    if tobecriture.Detail[i].Modifie then
    begin
      {JP 13/06/06 : FQ 18329 : pour mettre à jour la contrepartie en Tréso (ce qui signifie un changement de rubrique)}
      tobecriture.Detail[i].SetString('E_TRESOSYNCHRO', MajTreso(tobecriture.Detail[i], tobecriture.Detail[i].GetString('E_TRESOSYNCHRO') <> ets_Rien));
      //Récuperation de la tob des analytique
      CChargeAna(tobecriture.Detail[i]);
      //Mise à jour des contrepartie
      for j := 0 to tobecriture.Detail[i].Detail.Count - 1 do
      begin
        for z := 0 to tobecriture.Detail[i].Detail[j].Detail.Count - 1 do
        begin
          tobecriture.Detail[i].Detail[j].Detail[z].PutValue('Y_CONTREPARTIEAUX', tobecriture.Detail[i].GetValue('E_CONTREPARTIEAUX'));
          tobecriture.Detail[i].Detail[j].Detail[z].PutValue('Y_CONTREPARTIEGEN', tobecriture.Detail[i].GetValue('E_CONTREPARTIEGEN'));
        end;
      end;
    end;
  end;

  //Update dans la base
  tobecriture.UpdateDB;

  //Free da la tob
  FreeAndNil(tobecriture);

end;



procedure TOF_CPREIMPUT.InitMsgBox;
begin
  HM.Mess.Add('0;' + Ecran.Caption + ';Confirmez-vous le traitement de ré-imputation ?;Q;YN;Y;Y;');
  HM.Mess.Add('1;' + Ecran.Caption + ';Vous devez sélectionner au moins une écriture.;E;O;O;O;');
  HM.Mess.Add('2;' + Ecran.Caption + ';Vous devez sélectionner des écritures portant sur le même compte général.;W;O;O;O;');
  HM.Mess.Add('3;' + Ecran.Caption + ';Vous devez sélectionner des écritures portant sur le même compte auxiliaire.;W;O;O;O;');
  HM.Mess.Add('4;' + Ecran.Caption + ';Vous devez sélectionner des écritures portant sur le même établissement.;W;O;O;O;');
  HM.Mess.Add('5;' + Ecran.Caption + ';Vous devez sélectionner des écritures portant sur la même devise.;W;O;O;O;');
  HM.Mess.Add('6;' + Ecran.Caption + ';Vous devez sélectionner des écritures ayant le même taux de change.;W;O;O;O;');
  HM.Mess.Add('ATTENTION. Génération non effectuée ...');
  HM.Mess.Add('8;' + Ecran.Caption + ';Voulez-vous visualiser l''écriture générée ?;Q;YN;Y;Y;');
  HM.Mess.Add('EURO');
  HM.Mess.Add('10;' + Ecran.Caption + ';Voulez-vous visualiser l''écriture modifiée ?;Q;YN;Y;Y;');
  HM.Mess.Add('11;' + Ecran.Caption + ';Voulez-vous visualiser les écritures modifiées ?;Q;YN;Y;Y;');
  HM.Mess.Add('12;' + Ecran.Caption + ';Vous n''avez pas de sections d''attente sur l''axe ;W;O;O;O;');
end;

procedure TOF_CPREIMPUT.ChargeSelection;
var
  i: integer;
  M: RMVT;
  Q: TQuery;
begin
  InitMove(Fiche.FListe.NbSelected, '');
  VideListe(ListeSel);
  if Fiche.FListe.AllSelected then
  begin
    {$IFDEF EAGLCLIENT}
    Q := Fiche.Q.TQ;
(*    if not Fiche.FetchLesTous then
      PGIInfo('Impossible de récupérer tous les enregistrements.')
    else*)
      {$ELSE}
    Q := Fiche.Q;
    {$ENDIF}
    Q.First;
    while not Q.EOF do
    begin
      MoveCur(False);
      if TrouveSaisie(Q, M, GetControlText('E_QUALIFPIECE')) then
      begin
        M.NumLigne := Fiche.Q.FindField('E_NUMLIGNE').AsInteger;
        SaisComm.RemplirOFromM(M, ListeSel);
      end;
      Q.Next;
    end;
  end
  else
  begin
    for i := 0 to Fiche.FListe.NbSelected - 1 do
    begin
      MoveCur(False);
      Fiche.FListe.GotoLeBookMark(i);
      {$IFDEF EAGLCLIENT}
      Fiche.Q.TQ.Seek(Fiche.FListe.Row - 1);
      if TrouveSaisie(Fiche.Q.TQ, M, GetControlText('E_QUALIFPIECE')) then
      begin
        {$ELSE}
      if TrouveSaisie(Fiche.Q, M, GetControlText('E_QUALIFPIECE')) then
      begin
        {$ENDIF}
        M.NumLigne := Fiche.Q.FindField('E_NUMLIGNE').AsInteger;
        SaisComm.RemplirOFromM(M, ListeSel);
      end;
    end;
  end;
  FiniMove;
  Screen.Cursor := SyncrDefault;
end;

function TOF_CPREIMPUT.CoherenceOK: boolean;
var
  i, Err: integer;
  O: TOBM;
  OAux, OGene, OEtab, ODevise: string;
  OTaux: Double;
begin
  Err := 0;
  OTaux := 0;
  ChargeSelection;
  for i := 0 to ListeSel.Count - 1 do
  begin
    O := TOBM(ListeSel[i]);
    if i = 0 then
    begin
      OGene := O.GetMvt('E_GENERAL');
      OAux := O.GetMvt('E_AUXILIAIRE');
      OEtab := O.GetMvt('E_ETABLISSEMENT');
      ODevise := O.GetMvt('E_DEVISE');
      OTaux := O.GetMvt('E_TAUXDEV');
      if (ODevise <> V_PGI.DevisePivot) then gbDeviseNotPivot := True;
    end
    else
    begin
      if OGene <> O.GetMvt('E_GENERAL') then Err := 1 // Vous devez sélectionner des écritures portant sur le même compte général.
      else if OAux <> O.GetMvt('E_AUXILIAIRE') then Err := 2 // Vous devez sélectionner des écritures portant sur le même compte auxiliaire.
      else if OEtab <> O.GetMvt('E_ETABLISSEMENT') then Err := 3 // Vous devez sélectionner des écritures portant sur le même établissement.
      else if ODevise <> O.GetMvt('E_DEVISE') then Err := 4 // Vous devez sélectionner des écritures portant sur la même devise.
      else if OTaux <> O.GetMvt('E_TAUXDEV') then Err := 5; // Vous devez sélectionner des écritures ayant le même taux de change.
    end;
    if Err > 0 then Break;
  end;
  Result := (Err <= 0);
  if Err > 0 then HM.Execute(1 + Err, '', '');
end;

procedure TOF_CPREIMPUT.ContrePasse(TEcr: TOB);
begin
  TEcr.SetString('E_GENERAL', gszCONTRE_GENERAL);
  TEcr.SetString('E_AUXILIAIRE', gszCONTRE_AUXILIAIRE);
  if EncON then TEcr.SetString('E_TVAENCAISSEMENT', 'X')
           else TEcr.SetString('E_TVAENCAISSEMENT', '-');
end;

procedure TOF_CPREIMPUT.ContrePasseAna(TAna: TOB);
begin
  TAna.SetString('Y_GENERAL', gszCONTRE_GENERAL);
end;

// 14761
function TOF_CPREIMPUT.IsVentilable(General: string): Boolean;
var
  Q: TQuery;
begin
  Q := OpenSQL('SELECT G_VENTILABLE, G_VENTILABLE1, G_VENTILABLE2, G_VENTILABLE3, G_VENTILABLE4, G_VENTILABLE5 FROM GENERAUX WHERE G_GENERAL="' + General + '"', True);
  if Q.Eof then begin
    Result := False;
    FillChar(IsVenti, Sizeof(IsVenti), #0);
  end
  else begin
    Result := (Q.FindField('G_VENTILABLE').AsString = 'X');
    IsVenti[1] := (Q.FindField('G_VENTILABLE1').AsString = 'X');
    IsVenti[2] := (Q.FindField('G_VENTILABLE2').AsString = 'X');
    IsVenti[3] := (Q.FindField('G_VENTILABLE3').AsString = 'X');
    IsVenti[4] := (Q.FindField('G_VENTILABLE4').AsString = 'X');
    IsVenti[5] := (Q.FindField('G_VENTILABLE5').AsString = 'X');
  end;
  Ferme(Q);
end;

procedure TOF_CPREIMPUT.GenereLesPieces;
var
  NewM, OldM: RMVT;
  O1, O, OHisto: TOBM;
  i, NumL, k: integer;
  TEcr, TAna: TOB;
  Premier, OkAna, bVentilable: boolean;
  szAna: string;
  piecetraite, piecetraitefille: TOB;
begin
  FillChar(NewM, Sizeof(NewM), #0);
  FillChar(OldM, Sizeof(OldM), #0);
  FillChar(MsgSection, Sizeof(MsgSection), #0);
  VideListe(TPIECE);

  if gbDirect then
  begin
    // Modification directe de l'écriture
    for i := 0 to ListeSel.Count - 1 do
    begin
      MoveCur(False);
      O := TOBM(ListeSel[i]);
      OldM := OBMToIdent(O, False);
      NumL := O.GetValue('E_NUMLIGNE');
      bVentilable := IsVentilable(gszCONTRE_GENERAL);
      if bVentilable then szAna := 'E_ANA="X", ' // 14761
      else szAna := 'E_ANA="-", ';



      ExecuteSQL('UPDATE ECRITURE SET E_GENERAL="' + gszCONTRE_GENERAL + '", ' + szAna +
        {JP 22/08/05 : FQ 15931 : Destruction de l'éventuelle écriture de trésorerie et mise
                       à jour de E_TRESOSYNCHRO}
        'E_TRESOSYNCHRO = "' + MajTreso(O, True) + '", ' +
        'E_AUXILIAIRE="' + gszCONTRE_AUXILIAIRE + '", E_RIB="' + GetRibPrincipal(gszCONTRE_AUXILIAIRE) + '" WHERE ' +  //SG6 23/11/04 FQ 15006 Ajout du rib de l'auxiliaire
        'E_EXERCICE="' + O.GetValue('E_EXERCICE') + '" AND ' +
        'E_JOURNAL="' + O.GetValue('E_JOURNAL') + '" AND ' +
        'E_DATECOMPTABLE="' + USDATETIME(StrToDate(O.GetValue('E_DATECOMPTABLE'))) + '" AND ' +
        'E_NUMEROPIECE=' + IntToStr(O.GetValue('E_NUMEROPIECE')) + ' AND ' +
        'E_NUMLIGNE=' + IntToStr(NumL) + ' AND ' +
        'E_QUALIFPIECE="' + O.GetValue('E_QUALIFPIECE') + '" AND ' +
        'E_NUMECHE=' + IntToStr(O.GetValue('E_NUMECHE')));

      // Ajout me Fiche 19544 maj solde de l'ancien compte
      if O.GetValue ('E_GENERAL') <> '' then
          MajTotComptes([fbGene], FALSE, FALSE , O.GetValue ('E_EXERCICE') , O.GetValue ('E_GENERAL') );
      if O.GetValue ('E_AUXILIAIRE') <> '' then
          MajTotComptes([fbAux], FALSE, FALSE , O.GetValue ('E_EXERCICE') , O.GetValue ('E_AUXILIAIRE'));

      // Modification de ces valeurs pour la consultation
      O.PutValue('E_GENERAL', gszCONTRE_GENERAL);
      O.PutValue('E_AUXILIAIRE', gszCONTRE_AUXILIAIRE);
      TraiteLesAna(TAna, O, NumL, OldM, NewM, bVentilable);


      // Pour visualiser l'écriture générée
      OHisto := TOBM.Create(EcrGen, '', False);
      OHisto.Dupliquer(O, True, True);
      TPIECE.Add(OHisto);
    end;

    //Modification des E_ et Y_ CONTREPARTIE AUX OU GENE
    piecetraite := TOB.Create('TEMP', nil, -1);

    for i := 0 to ListeSel.Count - 1 do
    begin
      O := TOBM(ListeSel[i]);
      //On regarde si la piece a deja ete traité
      if piecetraite.FindFirst(['E_JOURNAL', 'E_EXERCICE', 'E_DATECOMPTABLE', 'E_NUMEROPIECE'], [O.GetValue('E_JOURNAL'), O.GetValue('E_EXERCICE'), O.GetValue('E_DATECOMPTABLE'),
        O.GetValue('E_NUMEROPIECE')], true) = nil then
      begin
        piecetraitefille := TOB.Create('TEMP', piecetraite, -1);
        // ajout me fiche 20161
        piecetraitefille.AddChampSupValeur('E_JOURNAL', O.GetValue('E_JOURNAL'));
        piecetraitefille.AddChampSupValeur('E_EXERCICE', O.GetValue('E_EXERCICE'));
        piecetraitefille.AddChampSupValeur('E_DATECOMPTABLE', O.GetValue('E_DATECOMPTABLE'));
        piecetraitefille.AddChampSupValeur('E_NUMEROPIECE', O.GetValue('E_NUMEROPIECE'));
        SetContrePartie(O);
      end;
    end;
    FreeAndNil(piecetraite);
  end
  else
  begin
    // Modification par OD
    O1 := TOBM(ListeSel[0]);
    OHisto := nil;

    // Nouvelle pièce
    NewM.Jal := gszCONTRE_JOURNAL;
    NewM.Exo := QuelExoDT(GeneDate);
    NewM.DateC := GeneDate;
    NewM.Num := GetNewNumJal(NewM.Jal, False, GeneDate);
    if NewM.Num <= 0 then
    begin
      V_PGI.IoError := oeUnknown;
      Exit;
    end;
    NewM.Simul := 'R';
    NewM.Nature := 'OD';
    NewM.CodeD := O1.GetMvt('E_DEVISE');
    NewM.TauxD := O1.GetMvt('E_TAUXDEV');
    NewM.DateTaux := GeneDate;
    NewM.Etabl := gszCONTRE_ETABLISSEMENT;

    TEcr := TOB.Create('ECRITURE', nil, -1);
    TAna := TOB.Create('ANALYTIQ', nil, -1);
    try
      // Ecritures
      InitMove(ListeSel.Count, '');
      NumL := -1;
      for i := 0 to ListeSel.Count - 1 do
      begin
        MoveCur(False);
        O := TOBM(ListeSel[i]);
        OldM := OBMToIdent(O, False);
        TripoteO(O, NewM, True);
        Premier := (O.GetMvt('E_NUMECHE') <= 1);

        // Ligne de ré-imputation
        if Premier then Inc(NumL, 2);
        O.SetCotation(0);
        O.SetMPACC;

        TEcr.Dupliquer(O, False, True, True);

        Reimpute(TEcr);
        TEcr.PutValue('E_NUMLIGNE', NumL);
        TEcr.PutValue('E_IO', 'X');

        {JP 22/08/05 : FQ 15931 : Mise à jour de E_TRESOSYNCHRO}
        TEcr.SetString('E_TRESOSYNCHRO', MajTreso(O, False));

        TraiteLesAna(TAna, O, NumL, OldM, NewM);

        TEcr.InsertOrUpdateDB(False);
        // Ligne de contrepassation
        OkAna := False;
        for k := 1 to MaxAxe do
          if Gen2Vent[k] then OkAna := True;

        TEcr.InitValeurs;

        O.SetCotation(0);
        O.SetMPACC;

        TEcr.Dupliquer(O, False, True, True);

        ContrePasse(TEcr);
        TEcr.PutValue('E_NUMLIGNE', NumL + 1);

        {JP 22/08/05 : FQ 15931 : Mise à jour de E_TRESOSYNCHRO}
        TEcr.SetString('E_TRESOSYNCHRO', MajTreso(O, False));

        if OkAna then TEcr.PutValue('E_ANA', 'X')
                 else TEcr.PutValue('E_ANA', '-');

        TEcr.InsertOrUpdateDB(False);

        // Pour visualiser l'écriture générée
        if ((Premier) and (OHisto = nil)) then
        begin
          OHisto := TOBM.Create(EcrGen, '', False);
          OHisto.Dupliquer(TEcr, True, True);
          TPIECE.Add(OHisto);
        end;
      end;
    finally
      FreeAndNil(TEcr);
      FreeAndNil(TAna);
    end;
  end;

  ADevalider(NewM.Jal, GeneDate);
  MarquerPublifi(True);
  // Ajout me Fiche 19544
  if gbDirect then MajdesSoldesImp;
  FiniMove;
end;

function TOF_CPREIMPUT.JePeuxValider: boolean;
begin
  Result := False;
  if (Fiche.FListe.NbSelected <= 0) and (not Fiche.FListe.AllSelected) then
  begin
    HM.Execute(1, '', '');
    Exit;
  end; // Vous devez sélectionner au moins une écriture.
  if not CoherenceOK then Exit;
  Result := True;
end;

procedure TOF_CPREIMPUT.Reimpute(TEcr: TOB);
var
  X: Double;
begin
  if gbContre_Negatif then
  begin
    TEcr.PutValue('E_DEBIT', -TEcr.GetValue('E_DEBIT'));
    TEcr.PutValue('E_CREDIT', -TEcr.GetValue('E_CREDIT'));
    TEcr.PutValue('E_DEBITDEV', -TEcr.GetValue('E_DEBITDEV'));
    TEcr.PutValue('E_CREDITDEV', -TEcr.GetValue('E_CREDITDEV'));
    TEcr.PutValue('E_QTE1', -TEcr.GetValue('E_QTE1'));
    TEcr.PutValue('E_QTE2', -TEcr.GetValue('E_QTE2'));
  end
  else
  begin
    X := TEcr.GetValue('E_DEBIT');
    TEcr.PutValue('E_DEBIT', TEcr.GetValue('E_CREDIT'));
    TEcr.PutValue('E_CREDIT', X);

    X := TEcr.GetValue('E_DEBITDEV');
    TEcr.PutValue('E_DEBITDEV', TEcr.GetValue('E_CREDITDEV'));
    TEcr.PutValue('E_CREDITDEV', X);
  end;
end;

procedure TOF_CPREIMPUT.ReimputeTobAna(Ana: Tob; OEcr: TOBM; Axe: Integer);
var
  szAxe, szSection: string;
begin
  szAxe := 'A' + IntToStr(Axe);
  szSection := VH^.Cpta[AxeToFb(szAxe)].Attente;

  if (szSection = '') and not MsgSection[Axe] then
  begin
    HM.Execute(12, '', IntToStr(Axe)); // Vous n''avez pas de sections d''attente sur l''axe
    MsgSection[Axe] := True; // Pour ne pas afficher plusieurs fois le message
    exit;
  end;
  Ana.PutValue('Y_GENERAL', gszCONTRE_GENERAL);
  Ana.PutValue('Y_AXE', szAxe);
  Ana.PutValue('Y_DATECOMPTABLE', OEcr.GetValue('E_DATECOMPTABLE'));
  Ana.PutValue('Y_NUMEROPIECE', OEcr.GetValue('E_NUMEROPIECE'));
  Ana.PutValue('Y_SECTION', szSection);
  Ana.PutValue('Y_EXERCICE', OEcr.GetValue('E_EXERCICE'));
  Ana.PutValue('Y_REFINTERNE', OEcr.GetValue('E_REFINTERNE'));
  Ana.PutValue('Y_LIBELLE', OEcr.GetValue('E_LIBELLE'));
  Ana.PutValue('Y_NATUREPIECE', OEcr.GetValue('E_NATUREPIECE'));
  Ana.PutValue('Y_QUALIFPIECE', OEcr.GetValue('E_QUALIFPIECE'));
  Ana.PutValue('Y_TYPEANALYTIQUE', '-');
  Ana.PutValue('Y_VALIDE', OEcr.GetValue('E_VALIDE'));
  Ana.PutValue('Y_ETAT', OEcr.GetValue('E_ETAT'));
  Ana.PutValue('Y_REFEXTERNE', OEcr.GetValue('E_REFEXTERNE'));
  Ana.PutValue('Y_DATEREFEXTERNE', OEcr.GetValue('E_DATEREFEXTERNE'));
  Ana.PutValue('Y_UTILISATEUR', OEcr.GetValue('E_UTILISATEUR'));
  Ana.PutValue('Y_CONTROLEUR', OEcr.GetValue('E_CONTROLEUR'));
  Ana.PutValue('Y_SOCIETE', OEcr.GetValue('E_SOCIETE'));
  Ana.PutValue('Y_ETABLISSEMENT', OEcr.GetValue('E_ETABLISSEMENT'));
  Ana.PutValue('Y_VISION', '-');
  Ana.PutValue('Y_AFFAIRE', OEcr.GetValue('E_AFFAIRE'));
  Ana.PutValue('Y_REFLIBRE', OEcr.GetValue('E_REFLIBRE'));
  Ana.PutValue('Y_DEVISE', OEcr.GetValue('E_DEVISE'));
  Ana.PutValue('Y_TAUXDEV', OEcr.GetValue('E_TAUXDEV'));
  Ana.PutValue('Y_DATETAUXDEV', OEcr.GetValue('E_DATETAUXDEV'));
  Ana.PutValue('Y_POURCENTAGE', '100');
  if Ana.GetValue('Y_DEBIT') <> 0 then Ana.PutValue('Y_TOTALECRITURE', Ana.GetValue('Y_DEBIT'))
  else Ana.PutValue('Y_TOTALECRITURE', Ana.GetValue('Y_CREDIT'));
  if Ana.GetValue('Y_DEBITDEV') <> 0 then Ana.PutValue('Y_TOTALECRITURE', Ana.GetValue('Y_DEBITDEV'))
  else Ana.PutValue('Y_TOTALECRITURE', Ana.GetValue('Y_CREDITDEV'));
  Ana.PutValue('Y_CONTROLE', OEcr.GetValue('E_CONTROLE'));
  Ana.PutValue('Y_TRACE', OEcr.GetValue('E_TRACE'));
  Ana.PutValue('Y_QTE1', OEcr.GetValue('E_QTE1'));
  Ana.PutValue('Y_QTE2', OEcr.GetValue('E_QTE2'));
  Ana.PutValue('Y_QUALIFQTE1', OEcr.GetValue('E_QUALIFQTE1'));
  Ana.PutValue('Y_QUALIFQTE2', OEcr.GetValue('E_QUALIFQTE2'));
  Ana.PutValue('Y_JOURNAL', OEcr.GetValue('E_JOURNAL'));
  Ana.PutValue('Y_NUMVENTIL', '1');
  Ana.PutValue('Y_TOTALQTE1', OEcr.GetValue('E_QTE1'));
  Ana.PutValue('Y_TOTALQTE2', OEcr.GetValue('E_QTE2'));
  Ana.PutValue('Y_TYPEANOUVEAU', OEcr.GetValue('E_TYPEANOUVEAU'));
  Ana.PutValue('Y_TYPEMVT', OEcr.GetValue('E_TYPEMVT'));
  Ana.PutValue('Y_ECRANOUVEAU', OEcr.GetValue('E_ECRANOUVEAU'));
  Ana.PutValue('Y_CONFIDENTIEL', OEcr.GetValue('E_CONFIDENTIEL'));
  Ana.PutValue('Y_CREERPAR', OEcr.GetValue('E_CREERPAR'));

  Ana.PutValue('Y_PERIODE', OEcr.GetValue('E_PERIODE'));
  Ana.PutValue('Y_SEMAINE', OEcr.GetValue('E_SEMAINE'));

  if gbContre_Negatif then
  begin
    Ana.PutValue('Y_DEBIT', -OEcr.GetValue('E_DEBIT'));
    Ana.PutValue('Y_CREDIT', -OEcr.GetValue('E_CREDIT'));
    Ana.PutValue('Y_DEBITDEV', -OEcr.GetValue('E_DEBITDEV'));
    Ana.PutValue('Y_CREDITDEV', -OEcr.GetValue('E_CREDITDEV'));
  end
  else
  begin
    //SG6 27.01.05 Correction effectuée (y avait inversion des credit debit)
    Ana.PutValue('Y_DEBIT', OEcr.GetValue('E_DEBIT'));
    Ana.PutValue('Y_CREDIT', OEcr.GetValue('E_CREDIT'));
    Ana.PutValue('Y_DEBITDEV', OEcr.GetValue('E_DEBITDEV'));
    Ana.PutValue('Y_CREDITDEV', OEcr.GetValue('E_CREDITDEV'));
  end;
end;

procedure TOF_CPREIMPUT.ReimputeAna(TAna: TOB);
var
  X: Double;
begin
  if gbContre_Negatif then
  begin
    TAna.PutValue('Y_DEBIT', -TAna.GetValue('Y_DEBIT'));
    TAna.PutValue('Y_CREDIT', -TAna.GetValue('Y_CREDIT'));
    TAna.PutValue('Y_DEBITDEV', -TAna.GetValue('Y_DEBITDEV'));
    TAna.PutValue('Y_CREDITDEV', -TAna.GetValue('Y_CREDITDEV'));
    TAna.PutValue('Y_QTE1', -TAna.GetValue('Y_QTE1'));
    TAna.PutValue('Y_QTE2', -TAna.GetValue('Y_QTE2'));
    TAna.PutValue('Y_TOTALQTE1', -TAna.GetValue('Y_TOTALQTE1'));
    TAna.PutValue('Y_TOTALQTE2', -TAna.GetValue('Y_TOTALQTE2'));
    TAna.PutValue('Y_TOTALECRITURE', -TAna.GetValue('Y_TOTALECRITURE'));
    TAna.PutValue('Y_TOTALDEVISE', -TAna.GetValue('Y_TOTALDEVISE'));
  end
  else
  begin
    X := TAna.GetValue('Y_DEBIT');
    TAna.PutValue('Y_DEBIT', TAna.GetValue('Y_CREDIT'));
    TAna.PutValue('Y_CREDIT', X);

    X := TAna.GetValue('Y_DEBITDEV');
    TAna.PutValue('Y_DEBITDEV', TAna.GetValue('Y_CREDITDEV'));
    TAna.PutValue('Y_CREDITDEV', X);
  end;
end;

procedure TOF_CPREIMPUT.TraiteLesAna(TAna: TOB; OEcr: TOBM; NewNumL: integer; OldM, NewM: RMVT; Ventilable: Boolean = True);
var
  QAna: TQuery;
  NumLE, i, k: integer;
  OA: TOBM;
  Ax, szSQL: string;
  DEV: RDEVISE;
  TobAna: Tob;
  iAxe: Integer;
  AVentiler: array[1..MaxAxe] of boolean;
begin
  VideListe(ListeAna);
  NumLE := OEcr.GetMvt('E_NUMLIGNE');

  QAna := OpenSQL('SELECT * FROM ANALYTIQ WHERE ' + WhereEcriture(tsAnal, OldM, False) + ' AND Y_NUMLIGNE=' + InttoStr(NumLE), True);
  while not QAna.EOF do
  begin
    Ax := QAna.FindField('Y_AXE').AsString;
    OA := TOBM.Create(EcrAna, Ax, False);
    OA.ChargeMvt(QAna);
    ListeAna.Add(OA);
    QAna.Next;
  end;
  Ferme(QAna);

  if gbDirect then
  begin
    // On parcourt la liste des écritures analytiques
    // 1° traitement :
    //  - Si le nouveau compte est ventilable sur un axe qui était précédemment ventilé : On modifie l'écriture
    //  - Si le nouveau compte n'est pas ventilable sur un axe qui était précédemment ventilé : On supprime l'écriture
    FillChar(AVentiler, Sizeof(AVentiler), #49);

    for i := 0 to ListeAna.Count - 1 do
    begin
      OA := TOBM(ListeAna[i]);
      Ax := OA.GetValue('Y_AXE');
      iAxe := StrToInt(Ax[2]);
      AVentiler[iAxe] := False;

      // Réimputation analytique
      if IsVenti[iAxe] then szSQL := 'UPDATE ANALYTIQ SET Y_GENERAL="' + gszCONTRE_GENERAL + '" WHERE ' // 14761
      else szSQL := 'DELETE FROM ANALYTIQ WHERE ';
      ExecuteSQL(szSQL +
        'Y_AXE="' + OA.GetValue('Y_AXE') + '" AND ' +
        'Y_DATECOMPTABLE="' + USDATETIME(StrToDate(OA.GetValue('Y_DATECOMPTABLE'))) + '" AND ' +
        'Y_NUMEROPIECE=' + IntToStr(OA.GetValue('Y_NUMEROPIECE')) + ' AND ' +
        'Y_NUMLIGNE=' + IntToStr(OA.GetValue('Y_NUMLIGNE')) + ' AND ' +
        'Y_EXERCICE="' + OA.GetValue('Y_EXERCICE') + '" AND ' +
        'Y_QUALIFPIECE="' + OA.GetValue('Y_QUALIFPIECE') + '" AND ' +
        'Y_JOURNAL="' + OA.GetValue('Y_JOURNAL') + '" AND ' +
        'Y_NUMVENTIL=' + IntToStr(OA.GetValue('Y_NUMVENTIL')));

         // Ajout me Fiche 19544
      MajTotComptes([fbsect], FALSE, FALSE , OA.GetValue('Y_EXERCICE'), OA.GetValue('Y_SECTION'), FALSE);

    end;

    // 2° traitement
    // Parcourt tous les autres axes et créé les écritures analytiques si nécesaire
    if not (VH^.AnaCroisaxe) then
    begin
      TobAna := TOB.Create('ANALYTIQ', nil, -1);
      for i := 1 to MaxAxe do
      begin
        if AVentiler[i] and IsVenti[i] then
        begin
          ReimputeTobAna(TobAna, OEcr, i);
          TobAna.PutValue('Y_NUMLIGNE', NewNumL);
          TobAna.InsertOrUpdateDB(False);
          // Ajout me Fiche 19544
          MajTotComptes([fbsect], FALSE, FALSE , TobAna.GetValue('Y_EXERCICE'), TobAna.GetValue('Y_SECTION'), FALSE);
        end;
      end;
    TobAna.Free;
    end;
  end
  else
  begin
    for i := 0 to ListeAna.Count - 1 do
    begin
      OA := TOBM(ListeAna[i]);
      TripoteO(OA, NewM, False);

      // Anal reimpute
      //TAna.Insert('ANALYTIQ');

      //OA.EgalChamps(TAna);
      TAna.Dupliquer(OA, False, True);

      ReimputeAna(TAna);
      TAna.PutValue('Y_NUMLIGNE', NewNumL);
      TAna.InsertOrUpdateDB(False);

      if MemeVentil then
      begin
        // Anal contrepasse
        //TAna.Insert('ANALYTIQ');

        //OA.EgalChamps(TAna.Detail[0]);
        TAna.Dupliquer(OA, False, True);

        ContrePasseAna(TAna);
        TAna.PutValue('Y_NUMLIGNE', NewNumL + 1);
        TAna.InsertOrUpdateDB(False);
      end;
    end;
    if MemeVentil then Exit;

    // Anal sur section attente
    FillChar(DEV, Sizeof(DEV), #0);
    VideListe(ListeAna);
    DEV.Code := OEcr.GetMvt('E_DEVISE');
    GetInfosDevise(DEV);
    DEV.DateTaux := GeneDate;
    DEV.Taux := OEcr.GetMvt('E_DATETAUXDEV');
    for i := 1 to MaxAxe do
    begin
      if Gen2Vent[i] then
      begin
        Ax := 'A' + Inttostr(i);
        VentileGenerale(Ax, OEcr, DEV, ListeAna, False);
        for k := 0 to ListeAna.Count - 1 do
        begin
          OA := TOBM(ListeAna[k]);
          TripoteO(OA, NewM, False);

          //OA.EgalChamps(TAna.Detail[0]);
          TAna.Dupliquer(OA, False, True);

          ContrePasseAna(TAna);
          TAna.PutValue('Y_NUMLIGNE', NewNumL + 1);
          TAna.InsertOrUpdateDB(False);
        end;
      end;
    end;
  end;
end;

procedure TOF_CPREIMPUT.TripoteO(O: TOBM; M: RMVT; EcrG: boolean);
begin
  if EcrG then
  begin
    O.PutMvt('E_JOURNAL', M.Jal);
    O.PutMvt('E_DATECREATION', Date);
    O.PutMvt('E_DATEMODIF', NowFutur);
    O.PutMvt('E_EXERCICE', M.Exo);
    O.PutMvt('E_DATECOMPTABLE', M.DateC);
    {$IFNDEF SPEC302}
    O.PutMvt('E_PERIODE', GetPeriode(M.DateC));
    O.PutMvt('E_SEMAINE', NumSemaine(M.DateC));
    {$ENDIF}
    O.PutMvt('E_NUMEROPIECE', M.Num);
    O.PutMvt('E_QUALIFPIECE', M.Simul);
    O.PutMvt('E_NATUREPIECE', M.Nature);
    O.PutMvt('E_DEVISE', M.CodeD);
    O.PutMvt('E_TAUXDEV', M.TauxD);
    O.PutMvt('E_DATETAUXDEV', M.DateTaux);
    O.PutMvt('E_ETABLISSEMENT', M.Etabl);
    O.PutMvt('E_QUALIFORIGINE', 'RIM');
    O.PutMvt('E_REFRELEVE', '');
    O.PutMvt('E_FLAGECR', '');
    O.PutMvt('E_TRACE', '');
    O.PutMvt('E_ETAT', '0000000000');
    O.PutMvt('E_NIVEAURELANCE', 0);
    O.PutMvt('E_DATERELANCE', IDate1900);
    O.PutMvt('E_VALIDE', '-');
    O.PutMvt('E_SUIVDEC', '');
    O.PutMvt('E_NOMLOT', '');
    O.PutMvt('E_FLAGECR', '');
    O.PutMvt('E_CONTROLETVA', '');
    O.PutMvt('E_EDITEETATTVA', '-');
    O.PutMvt('E_NIVEAURELANCE', 0);
  end
  else
  begin
    O.PutMvt('Y_JOURNAL', M.Jal);
    O.PutMvt('Y_DATECREATION', Date);
    O.PutMvt('Y_DATEMODIF', NowFutur);
    O.PutMvt('Y_EXERCICE', M.Exo);
    O.PutMvt('Y_DATECOMPTABLE', M.DateC);
    {$IFNDEF SPEC302}
    O.PutMvt('Y_PERIODE', GetPeriode(M.DateC));
    O.PutMvt('Y_SEMAINE', NumSemaine(M.DateC));
    {$ENDIF}
    O.PutMvt('Y_NUMEROPIECE', M.Num);
    O.PutMvt('Y_QUALIFPIECE', M.Simul);
    O.PutMvt('Y_NATUREPIECE', M.Nature);
    O.PutMvt('Y_DEVISE', M.CodeD);
    O.PutMvt('Y_TAUXDEV', M.TauxD);
    O.PutMvt('Y_DATETAUXDEV', M.DateTaux);
    O.PutMvt('Y_ETABLISSEMENT', M.Etabl);
    O.PutMvt('Y_TRACE', '');
    O.PutMvt('Y_VALIDE', '-');
  end;
end;

procedure TOF_CPREIMPUT.E_EXERCICEChange(Sender: TObject);
begin
  if THValComboBox(Sender).Text <> '<<Tous>>' then //SG6 15/11/04 FQ 14959
  begin
    ExoToDates(GetControlText('E_EXERCICE'), GetControl('E_DATECOMPTABLE', True), GetControl('E_DATECOMPTABLE_', True));
  end
  else
  begin
    SetControlText('E_DATECOMPTABLE', DateToStr(VH^.EnCours.Deb)); //SG6 15/11/04 FQ 14959
    SetControlText('E_DATECOMPTABLE_', DateToStr(VH^.Suivant.Fin));  //SG6 15/11/04 FQ 14959
  end;

end;

procedure TOF_CPREIMPUT.BListePIECESClick(Sender: TObject);
begin
  if TPiece.Count > 0 then VisuPiecesGenere(TPiece, EcrGen, 8);
end;

procedure TOF_CPREIMPUT.FListeDblClick(Sender: TObject);

begin
  {$IFDEF EAGLCLIENT}
  Fiche.Q.TQ.Seek(Fiche.FListe.Row - 1);
  TrouveEtLanceSaisie(Fiche.Q.TQ, taConsult, GetControlText('E_QUALIFPIECE'));
  {$ELSE}
  TrouveEtLanceSaisie(Fiche.Q, taConsult, GetControlText('E_QUALIFPIECE'));
  {$ENDIF}
end;

procedure TOF_CPREIMPUT.BZoomClick(Sender: TObject);
begin
  FListeDBLClick(nil);
end;

procedure TOF_CPREIMPUT.BChercheClick(Sender: TObject);
begin
  VideListe(ListeSel);
  Fiche.BChercheClick(Sender); // inherited;
  gbDeviseNotPivot := False;
  {$IFDEF EAGLCLIENT}
  GereSelectionsGrid(Fiche.FListe, Fiche.Q.TQ);
  {$ELSE}
  GereSelectionsGrid(Fiche.FListe, Fiche.Q);
  {$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 13/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPREIMPUT.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;

procedure TOF_CPREIMPUT.BOuvrirClick(Sender: TObject);
var
  TInfo: TOB;
  O1: TOBM;
  i: Integer;
begin
  MemeVentil := False;
  if not JePeuxValider then
  begin
    Screen.Cursor := SyncrDefault;
    Exit;
  end;

  TInfo := TOB.Create('', nil, -1);
  O1 := TOBM(ListeSel[0]);


  // Pour l'envoi des paramètres
  if gbDeviseNotPivot then TInfo.AddChampSupValeur('GBDEVISENOTPIVOT', 'X') // CoherenceOK
  else TInfo.AddChampSupValeur('GBDEVISENOTPIVOT', '-');
  TInfo.AddChampSupValeur('GSZAUXILIAIRE', O1.GetMvt('E_AUXILIAIRE')); // utilisé dans ComptesOK
  TInfo.AddChampSupValeur('GSZGENERAL', O1.GetMvt('E_GENERAL')); // utilisé dans ComptesOK
  TInfo.AddChampSupValeur('GSZDEVISE', O1.GetMvt('E_DEVISE')); // utilisé dans ComptesOK

  {JP 01/08/05 : FQ 15124 : On regarde s'il y a des lignes de TTC Gescom dans la sélection}
  if PresenceTTCGescom then begin
    {Il y a des TTC Gescom, on n'autorise pas la réimputation directe}
    TInfo.AddChampSupValeur('CANDIRECT', '-');
    {On avertit l'utilisateur}
    HShowMessage('0;' + Ecran.Caption + ';Il y a des écritures de TTC qui ne peuvent être modifiées.'#13 +
                                         'La modification directe des écritures n''est pas possible,'#13 +
                                         'seule la création d''écritures de réimputation le sera.;I;O;O;O;', '', '');
  end

  {JP 26/06/07 : FQ TRESO 10491 : on regarde s'il y a des flux originaires de la Trésorerie}
  else if PresenceMvtTreso then begin
    {Il y a des flux de Tréso, on n'autorise pas la réimputation directe}
    TInfo.AddChampSupValeur('CANDIRECT', '-');
    {On avertit l'utilisateur}
    HShowMessage('0;' + Ecran.Caption + ';Il y a des écritures issues de la Trésorerie qui ne peuvent être modifiées.'#13 +
                                         'La modification directe des écritures n''est pas possible,'#13 +
                                         'seule la création d''écritures de réimputation le sera.;I;O;O;O;', '', '');
  end
  else
    {Il n'y a pas de TTC Gescom, on autorise la réimputation directe}
    TInfo.AddChampSupValeur('CANDIRECT', 'X');

  TheTOB := TInfo;


  CC_LanceFicheReimpParam;
  TInfo.Free;

  if TheTOB = nil then
  begin
    Exit;
  end;
  // Récupère les paramètres
  gbDirect := (TheTOB.GetValue('DIRECT') = 'X');
  gszCONTRE_GENERAL := TheTOB.GetValue('CONTRE_GENERAL');
  gszCONTRE_AUXILIAIRE := TheTOB.GetValue('CONTRE_AUXILIAIRE');
  gszCONTRE_JOURNAL := TheTOB.GetValue('CONTRE_JOURNAL');
  gszCONTRE_ETABLISSEMENT := TheTOB.GetValue('CONTRE_ETABLISSEMENT');
  if TheTOB.GetValue('CONTRE_DATE') <> '' then
    GeneDate := StrToDate(TheTOB.GetValue('CONTRE_DATE'));
  gbCONTRE_NEGATIF := (TheTOB.GetValue('CONTRE_NEGATIF') = 'X');
  MemeVentil := TheTOB.GetValue('MEMEVENTIL');
  for i := 1 to MaxAxe do
  begin // SBO Correction FFF 127
    Gen2Vent[i] := TheTOB.GetValue('GEN2VENT' + IntToStr(i));
  end;

  // Confirmez-vous la génération de l'écriture de ré-imputation ?
  if HM.Execute(0, '', '') <> mrYes then Exit;

  Application.ProcessMessages;
  NowFutur := NowH;

  if Transactions(GenereLesPieces, 3) <> oeOK then
    MessageAlerte(HM.Mess[7]) // Attention. Génération non effectuée.
  else
  begin
    if TPiece.Count > 0 then
      if gbDirect then
      begin
        if ListeSel.Count = 1 then
        begin
          // Voulez-vous visualiser l'écriture modifiée ?
          if HM.Execute(10, '', '') = mrYes then VisuPiecesGenere(TPiece, EcrGen, 8);
        end
        else
        begin
          // Voulez-vous visualiser les écritures modifiées ?
          if HM.Execute(11, '', '') = mrYes then VisuPiecesGenere(TPiece, EcrGen, 8);
        end;
      end
      else
        // Voulez-vous visualiser l'écriture générée ?
        if HM.Execute(8, '', '') = mrYes then VisuPiecesGenere(TPiece, EcrGen, 8);
  end;
  TFMul(Ecran).BChercheClick(nil);
end;

{JP 01/08/05 : FA 15124 : On regarde s'il y a dans la sélection une ligne TTC de Gescom
{---------------------------------------------------------------------------------------}
function TOF_CPREIMPUT.PresenceTTCGescom : Boolean;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  Result := False;
  // 51320 : Specif TESSI TVA : possibilité de réimputation sur pièce GESCOM + onglet avancé (CPREIMPUT_TOF.PAS)
  if EstSpecif('51320') then Exit ;
  for n := 0 to TFMul(Ecran).FListe.NbSelected - 1 do begin
    TFMul(Ecran).FListe.GotoLeBookmark(n);
    if (VarToStr(GetField('E_REFGESCOM')) <> '') and
       (VarToStr(GetField('E_TYPEMVT')) = 'TTC') then begin
      Result := True;
      Break;
    end;
  end;
end;

{JP 26/06/07 : FQ TRESO 10491 : on empêche la suppression des écritures issues de la Trésorerie
{---------------------------------------------------------------------------------------}
function TOF_CPREIMPUT.PresenceMvtTreso : Boolean;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  Result := False;
  {On ne teste que si la Tréso est utilisée sur cette base}
  if not EstComptaTreso then Exit;

  for n := 0 to TFMul(Ecran).FListe.NbSelected - 1 do begin
    TFMul(Ecran).FListe.GotoLeBookmark(n);
    if (VarToStr(GetField('E_QUALIFORIGINE')) = QUALIFTRESO) then begin
      Result := True;
      Break;
    end;
  end;
end;

{JP 22/08/05 : FQ 15931 : Suppression des écritures en Treso et maj de E_TRESOSYNCHRO}
{---------------------------------------------------------------------------------------}
function TOF_CPREIMPUT.MajTreso(T : TOB; Detruire : Boolean) : string;
{---------------------------------------------------------------------------------------}
begin
  {Destruction éventuelle de l'écriture en tréso}
  if Detruire then
    DetruitEcritureTresoTOB(T);

  {Recherche de la valeur à affecter à E_TRESOSYNCHRO}
  //if EstCptSyn(T.GetString('E_GENERAL'), T.GetString('E_JOURNAL'), T.GetString('E_NATUREPIECE') <> 'ECC') then
  if T.GetString('E_TRESOSYNCHRO') = ets_Rien then
    Result := ets_Rien
  else
    Result := ets_Nouveau;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPREIMPUT.SlctAllClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Fiche : TFMul;
begin
  Fiche := TFMul(Ecran);
  {$IFDEF EAGLCLIENT}
  if not Fiche.FListe.AllSelected then begin
    if not Fiche.FetchLesTous then Exit;
  end;
  {$ENDIF}
  Fiche.bSelectAllClick(nil);
end;

      // Ajout me Fiche 19544
procedure TOF_CPREIMPUT.MajdesSoldesImp;
var
i : integer;
O : TOB;
begin
   for i:=0 to TPiece.Count-1 do
   begin
       O:=TOB(TPiece[i]) ;
       if O.GetValue ('E_GENERAL') <> '' then
          MajTotComptes([fbGene], FALSE, FALSE , O.GetValue ('E_EXERCICE') , O.GetValue ('E_GENERAL') , FALSE);
       if O.GetValue ('E_AUXILIAIRE') <> '' then
          MajTotComptes([fbAux], FALSE, FALSE , O.GetValue ('E_EXERCICE') , O.GetValue ('E_AUXILIAIRE') , FALSE);
       MajTotComptes([fbJal], FALSE, FALSE , O.GetValue ('E_EXERCICE'), O.GetValue ('E_JOURNAL'), FALSE);
   end;
end;

initialization
  registerclasses([TOF_CPREIMPUT]);
end.

