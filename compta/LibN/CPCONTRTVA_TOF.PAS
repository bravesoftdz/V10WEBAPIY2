{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 04/03/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPCONTRTVA ()
Mots clefs ... : TOF;CPCONTRTVA
*****************************************************************}
Unit CPCONTRTVA_TOF ;

Interface

Uses Controls,     // mrYe, mrNo, mrCancel
     Forms,        // Application
     Classes,      // registerclasses
     Menus,        // TMenuItem
     ComCtrls,     // TPageControl
     SysUtils,     // DateToStr
     UtilPGI,      // _Blocage
     ed_tools,     // VideListe
{$IFDEF EAGLCLIENT}
     MaineAGL,     // AGLLanceFiche
     eMul,         // TFMul
     UTob,         // TQuery
     CPTVATPF_TOF, // ParamTVATPF
{$ELSE}
     FE_Main,      // AGLLanceFiche
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Mul,          // TFMul
     CPTVATPF_TOF,          // ParamTVATPF
     HDB,          // HDBGrid
{$endIF}
     Saisie,       // TrouveSaisie, TrouveEtLanceSaisie
     Ent1,         // ExoToDates, VH, Tva2Taux,
     HCtrls,       // THValComboBox
     HStatus,      // InitMove, MoveCur, FiniMove
     HCompte,      // TGGeneral, TGTiers
     HEnt1,        // Valeur, Arrondi,
     SaisComm,     // RenseigneHTByTva,
     SaisUtil,     // GereSelectionsGrid
     HTB97,        // TToolBarButton97
     HMsgBox,      // THMsgBox
     uLibWindows,
     UTOF,
     ParamSoc,		 // GetParamSocSecur YMO
     StdCtrls;

procedure ControleTVAFactures ( EXO : String3 ; D1,D2 : TDateTime ; Etab : String ) ;
procedure RecupInfosTVA ( TEcr,TGen,TAux : TList ; Var RegTVA : String3 ; Var SoumisTPF : boolean ) ;
procedure ConstitueListesTVA ( M : RMVT ; TEcr,TGen,TAux : TList ; Var Nbl : integer ) ;
function  EstTiersGeneTVA ( CGen : TGGeneral ) : boolean ;
function  NatureHTTVA ( NatG : String ) : boolean ;

type X_TVA = Class
             CHT,CTVA,CTPF  : String17 ;
             RegTVA,CodeTva : String3 ;
             Achat          : boolean ;
             XHTP,XTVAAP,XTVADP,XTPFAP,XTPFDP : Double ;
             XHTD,XTVAAD,XTVADD,XTPFAD,XTPFDD : Double ;
//             XHTE,XTVAAE,XTVADE,XTPFAE,XTPFDE : Double ;
end;

Type
  TOF_CPCONTRTVA = Class (TOF)
{$IFDEF EAGLCLIENT}
    FListe : THGrid;
{$ELSE}
    FListe : THDBGrid;
{$ENDIF}
    HM : THMsgBox;
    BParamTPF  : TMenuItem;
    BParamTVA  : TMenuItem;
    BZoomPiece : TMenuItem;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure InitMsgBox;

    procedure BOuvrirClick(Sender: TObject);
    procedure BChercheClick(Sender: TObject);

    procedure BZoomPieceClick(Sender: TObject);
    procedure BParamTVAClick(Sender: TObject);
    procedure BParamTPFClick(Sender: TObject);
    procedure BSelectAllClick(Sender: TObject);
    procedure E_EXERCICEChange(Sender: TObject);
    procedure E_ETABLISSEMENTEnter(Sender: TObject);
    procedure TVAOKOnClick(sender: TObject);

    procedure FListeDblClick(Sender: TObject);

    procedure AuxiElipsisClick ( Sender : TObject );
  private
    NowFutur   : TDateTime ;
    TT_TVA     : TList ;
    DEV        : RDEVISE ;
    TVAOK : THCheckbox ;
    procedure ControleLaTva ;
    function  ControlePiece ( M : RMVT ; Var NbL : integer ) : String ;
    procedure MajEtat ( M : RMVT ; EtatTVA : String3 ; NbL : integer ) ;
// Conversions, calculs
    function  PasDeCodeTVA ( TGen,TEcr : TList ) : boolean ;
// Etude de la pièce
    function  EtudieTVA ( TEcr,TGen,TAux : TList ; Achat, OD : boolean ) : String3 ;
    function  PasDeTiers ( TGen,TAux : TList ) : boolean;
    function  PasDeHT ( TGen : TList ) : boolean;
    function  MultiTiers ( TEcr,TGen,TAux : TList ) : boolean;
    function  TiersSansRegime ( TGen,TAux : TList ) : boolean;
    procedure RempliTVA ( TEcr,TGen : TList; RegTVA : String3; Achat,SoumisTPF : boolean );
    procedure CumulTVA ( XX : X_TVA; O : TOBM; CGen : TGGeneral; RegTVA : String3; Achat,SoumisTPF : boolean );
    function  SommeSais ( CPT : String; TEcr : TList; tsm : TSorteMontant ) : Double;
    function  SommeCalc ( CPT : String; OkTVA,Arr : Boolean; tsm : TSorteMontant ) : Double;
    function  PresentTVA ( CPT : String; TEcr : TList ) : boolean;
end ;

Implementation

uses
  {$IFDEF MODENT1}
  ULibExercice,
  CPProcMetier,
  {$ENDIF MODENT1}
  {$IFDEF eAGLCLIENT}
  MenuOLX
  {$ELSE}
  MenuOLG
  {$ENDIF eAGLCLIENT}
  , UTofMulParamGen; {13/04/07 YMO F5 sur Auxiliaire }


procedure ControleTVAFactures ( EXO : String3 ; D1,D2 : TDateTime ; Etab : String ) ;
var
  szArg : String;
begin
  if _Blocage(['nrCloture','nrBatch'],False,'nrBatch') then Exit ;
  szArg := Exo + ';' + Etab + ';' + DateToStr(D1) + ';' + DateToStr(D2);
  AGLLanceFiche('CP','CPCONTRTVA','','',szArg);
end;

procedure RecupInfosTVA ( TEcr,TGen,TAux : TList ; Var RegTVA : String3 ; Var SoumisTPF : boolean ) ;
var
  i : integer ;
  CGen : TGGeneral ;
  CAux : TGTiers ;
  O    : TOBM ;
begin
  for i := 0 to TEcr.Count-1 do begin
    O := TOBM(TEcr[i]);
    if O=Nil then Continue ;
    if O.GetMvt('E_NUMECHE')<=1 then begin
      RegTva := O.GetMvt('E_REGIMETVA') ;
      CGen := TGGeneral(TGen[i]) ;
      if EstTiersGeneTVA(CGen) then begin
        if RegTva = '' then RegTVA := CGen.RegimeTVA ;
        SoumisTPF := CGen.SoumisTPF ;
        Break ;
        end
      else if TGTiers(TAux[i])<>Nil then begin
        CAux := TGTiers(TAux[i]) ;
        if RegTva = '' then RegTVA := CAux.RegimeTVA ;
        SoumisTPF := CAux.SoumisTPF ;
        Break ;
      end;
    end;
  end;
  if RegTva = '' then RegTva := VH^.RegimeDefaut ;
end;

procedure ConstitueListesTVA ( M : RMVT ; TEcr,TGen,TAux : TList ; Var Nbl : integer ) ;
var
  QEcr : TQuery ;
  O    : TOBM ;
  CGen : TGGeneral ;
  CAux : TGTiers ;
  stSQL : string ;
  CptGen,CptAux : String ;
begin
  Nbl:=0 ;
  StSQL := 'Select ' ;
  stSQL := stSQL + 'E_REGIMETVA,E_JOURNAL,E_EXERCICE,E_DATECOMPTABLE,';
  stSQL := stSQL + 'E_NUMEROPIECE,E_QUALIFPIECE,E_GENERAL,E_AUXILIAIRE,';
  stSQL := stSQL + 'E_NATUREPIECE,E_TVA,E_TVAENCAISSEMENT,E_MODESAISIE,E_VALIDE,';
  stSQL := stSQL + 'E_DEVISE,E_TAUXDEV,E_DATETAUXDEV,E_NUMLIGNE,E_NUMECHE,E_ECHE,';
  stSQL := stSQL + 'E_ECHEENC1,E_ECHEENC2,E_ECHEENC3,E_ECHEENC4,E_ECHEDEBIT,E_TPF,E_TVA,';
  stSQL := stSQL + 'E_ETATREVISION,E_QUALIFORIGINE,E_TVAENCAISSEMENT,E_EMETTEURTVA,E_EDITEETATTVA,';
  stSQL := stSQL + 'E_QUALIFORIGINE,E_DEBIT,E_CREDIT ' ;
 { YMO 19/05/2006&05/06/2006 FQ17811 E_DEBIT et E_CREDIT n'étaient pas ramenés, les calculs ne pouvaient alors se faire}
  stSQL := stSQL + 'from Ecriture Where '+ WhereEcriture(tsGene,M,False);
  try
    QEcr := OpenSQL(stSQL,True) ;
    While Not QEcr.EOF do begin
    // YMO 01/06/2006 FQ17811 En oracle version 2 tiers, le recordcount renvoie -1
    //for i:=0 to Qecr.RecordCount-1 do begin
      O := TOBM.Create(EcrGen,'',False) ;
      O.ChargeMvt(QEcr) ;
      CptGen := O.GetMvt('E_GENERAL') ;
      CGen := TGGeneral.Create(CptGen) ;
      CptAux := O.GetMvt('E_AUXILIAIRE') ;
      if CptAux<>'' then CAux := TGTiers.Create(CptAux)
                    else CAux:=Nil ;
      inc(NbL) ;
      TEcr.Add(O) ;
      TGen.Add(CGen) ;
      TAux.Add(CAux) ;
      //FreeAndNil(CGen);
      //if CptAux<>'' then FreeAndNil(CAux);
      //FreeAndNil(CAux);
      QEcr.Next ;
    end;
    Ferme(QEcr) ;
  except
    Ferme(QEcr) ;
  end ;
end;

function  EstTiersGeneTVA ( CGen : TGGeneral ) : boolean ;
begin
  Result := ((CGen.NatureGene='TID') or (CGen.NatureGene='TIC'));
end;

function  NatureHTTVA ( NatG : String ) : boolean ;
begin
  Result := ((NatG='CHA') or (NatG='PRO') or (NatG='IMO'));
end;

procedure TOF_CPCONTRTVA.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPCONTRTVA.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPCONTRTVA.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPCONTRTVA.OnLoad ;
begin
  Inherited ;
 (*
{$IFDEF EAGLCLIENT}
  GereSelectionsGrid(FListe,TFMul(Ecran).Q.TQ);
{$ELSE}
  GereSelectionsGrid(FListe,TFMul(Ecran).Q);
{$ENDIF}
*)
end ;

procedure TOF_CPCONTRTVA.OnArgument (S : String ) ;
var
  LeExo     : String3;
  LaD1,LaD2 : String;
  LesEtab   : String;
begin
  Inherited ;
  // Création des contrôles
  HM:= THMsgBox.create(FMenuG);
  InitMsgBox;
  // Evénéments des contrôles
  TVAOK := THCheckbox(GetControl('TVAOK', True)) ;
  TVAOK.OnClick := TVAOKOnClick ;

  THValComboBox(GetControl('E_EXERCICE',True)).OnChange := E_EXERCICEChange;
  TToolBarButton97(GetControl('BOUVRIR',True)).OnClick := BOuvrirClick;
  TToolBarButton97(GetControl('BCHERCHE',True)).OnClick := BChercheClick;
  TMenuItem(GetConTrol('BPARAMTPF',True)).OnClick  := BParamTPFClick;
  TMenuItem(GetConTrol('BPARAMTVA',True)).OnClick  := BParamTVAClick;
  TMenuItem(GetConTrol('BZOOMPIECE',True)).OnClick := BZoomPieceClick;
  TToolBarButton97(GetControl('BSELECTALL',True)).OnClick := BSelectAllClick;
  THMultiValComboBox(GetControl('E_ETABLISSEMENT',True)).OnEnter := E_ETABLISSEMENTEnter ;

{$IFDEF EAGLCLIENT}
  FListe := THGrid(GetControl('FLISTE',True));
{$ELSE}
  FListe := THDBGrid(GetControl('FLISTE',True));
{$ENDIF}
  FListe.OnDblClick := FListeDblClick;

  //SG6 03.03.05 Fiche Qualite 15436
  FListe.SortEnabled := True;

  // Récupère les arguments
  LeExo := ReadTokenSt(S);
  LaD1 := ReadTokenSt(S);
  LaD2 := ReadTokenSt(S);
  LesEtab := ReadTokenSt(S);

  //////////////////////////////////////////////////////////////////
  // FormCreate
  //TFMul(Ecran).Q.Manuel := False ; //True;
  //TFMul(Ecran).TypeAction := taConsult ; // taModif;
  // YM 22/09/2005 N'a plus lieu d'être ; n'était pris en compte qu'en 2 tiers
  //TFMul(Ecran).FNomFiltre := 'CONTROLETVA';

  SetControlText('E_QUALIFPIECE','N');
  SetControlText('E_JOURNAL', '');
  TT_TVA := TList.Create;

  //////////////////////////////////////////////////////////////////
  // FormShow
  if LeExo='' then begin
    if VH^.CPExoRef.Code<>'' then begin
      SetControlText('E_EXERCICE',VH^.CPExoRef.Code);
      SetControlText('E_DATECOMPTABLE',DateToStr(VH^.CPExoRef.Deb));
      SetControlText('E_DATECOMPTABLE_',DateToStr(VH^.CPExoRef.Fin));
      end
    else begin
      SetControlText('E_EXERCICE',VH^.Entree.Code);
      SetControlText('E_DATECOMPTABLE',DateToStr(V_PGI.DateEntree));
      SetControlText('E_DATECOMPTABLE_',DateToStr(V_PGI.DateEntree));
    end;
    end
  else begin
    SetControlText('E_EXERCICE',LeExo);
    SetControlText('E_DATECOMPTABLE',LaD1);
    SetControlText('E_DATECOMPTABLE_',LaD2);
    SetControlText('E_ETABLISSEMENT',LesEtab);
  end;

  // FQ13924 SBO : Forcer à <<Tous>> si chaîne vide afin de détecter le passage sur la zone
  if GetControlText('E_ETABLISSEMENT')='' then
    SetControlText('E_ETABLISSEMENT', TraduireMemoire('<<Tous>>') ) ;
  PositionneEtabUser(GetControl('E_ETABLISSEMENT'), False); // 15090

  //SetTousCombo(THValComboBox(GetControl('E_JOURNAL',True)));
  TPageControl(GetControl('PAGES',True)).Pages[3].TabVisible := False;
  //TFMul(Ecran).Q.Manuel := False;
  //TFMul(Ecran).Q.UpdateCriteres;


  if GetParamSocSecur('SO_CPMULTIERS', false) then
    THEdit(GetControl('E_AUXILIAIRE', true)).OnElipsisClick:=AuxiElipsisClick;

end ;

procedure TOF_CPCONTRTVA.OnClose ;
begin
  Inherited ;
  HM.Free;

  //////////////////////////////////////////////////////////////////
  // FormClose
  VideListe(TT_TVA);
  FreeAndNil(TT_TVA);
  if IsInside(Ecran) then _Bloqueur('nrBatch',False);
end ;

procedure TOF_CPCONTRTVA.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPCONTRTVA.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPCONTRTVA.E_EXERCICEChange(Sender: TObject);
begin
  ExoToDates(GetControlText('E_EXERCICE'),GetControl('E_DATECOMPTABLE',True),GetControl('E_DATECOMPTABLE_',True));
end;

procedure TOF_CPCONTRTVA.ControleLaTva;
var
  i,MM,maxM, Nbl : integer ;
  M : RMVT ;
  EtatTVA : String3 ;
  NbrE : Integer;
  All : Boolean ;
begin
  try
    MaxM := 10 ;
    All := FListe.AllSelected ;
    If All then
    begin
    {$IFDEF EAGLCLIENT}
        NbrE := TFMul(Ecran).Q.TQ.RecordCount ;
        TFMul(Ecran).Q.TQ.First;
    {$ELSE}
        NbrE := TFMul(Ecran).Q.RecordCount ;
        TFMul(Ecran).Q.First;
    {$ENDIF}
      //InitMove(round(NbrE/maxM)+1,'Contrôle TVA en cours...')
      InitMoveProgressForm(ecran,Ecran.caption,TraduireMemoire('Contrôle TVA en cours...'), NbrE, true, true);
    end
    else
    begin
      NbrE := FListe.nbSelected ;
      InitMove(round(NbrE/maxM)+1,'Contrôle TVA en cours...')
    end;
    MM := MaxM ;
    MoveCurProgressForm('Pièce traitée : ' + IntToStr(M.Num)) ;
    {$IFNDEF EAGLCLIENT}
    if All then  //Allselected
    while not
        TFMul(Ecran).Q.EOF
    do
    begin
        if TrouveSaisie(TFMul(Ecran).Q,M,GetControlText('E_QUALIFPIECE')) then begin
      if (MM=0) then
      begin
        MM := maxM ;
        MoveCurProgressForm('Pièce traitée : ' + IntToStr(M.Num)) ;    // on divise par 10 le nb de fois où l'on procède au refresh
      end ;
      if MM=maxM then MoveCurProgressForm('Pièce traitée : ' + IntToStr(M.Num)) ;
      dec(MM) ;
        EtatTVA := ControlePiece(M,NbL) ;
        MajEtat(M,EtatTVA,Nbl);
        FreeAndNil(M);
        end; // TrouveSaisie
        TFMul(Ecran).Q.Next;
    end
    else // ligne à ligne
    {$ENDIF}
    for i:=0 to TFMul(Ecran).FListe.NbSelected - 1 do //NbrE-1 do
    begin
      dec(MM) ;
      if ((MM=0) or (i=0)) then
      begin
        MoveCur(False) ;    // on divise par 10 le nb de fois où l'on procède au refresh
        MM := MaxM ;
      end ;
      FListe.GotoLeBookMark(i) ;
      {$IFDEF EAGLCLIENT}
        TFMul(Ecran).Q.TQ.Seek(FListe.Row-1);
        if TrouveSaisie(TFMul(Ecran).Q.TQ,M,GetControlText('E_QUALIFPIECE')) then begin
        {$ELSE}
        if TrouveSaisie(TFMul(Ecran).Q,M,GetControlText('E_QUALIFPIECE')) then begin
        {$ENDIF}
          EtatTVA := ControlePiece(M,NbL) ;
          MajEtat(M,EtatTVA,Nbl);
          FreeAndNil(M);
        end; // TrouveSaisie
    end; //For
  except
    FreeAndNil(M);
    MessageAlerte(HM.Mess[2]); // Certaines pièces en cours de traitement ont stoppé la validation.
  end ;
  //FiniMove ;
  FiniMoveProgressForm ;

end;

function TOF_CPCONTRTVA.ControlePiece(M: RMVT; var NbL: integer): String;
var
  QJal : TQuery ;
  TEcr,TGen,TAux : TList ;
  Etat      : String3 ;
  Achat, OD : boolean ;
begin
  DEV.Code := M.CodeD ;
  GetInfosDevise(DEV) ;
  Etat:='COR' ;
  QJal:=OpenSQL('SELECT J_NATUREJAL FROM JOURNAL WHERE J_JOURNAL="'+M.Jal+'"',True) ;
  if Not QJal.EOF then Achat := (QJal.Fields[0].AsString='ACH')
                    else Achat:=False ;
    OD := (QJal.Fields[0].AsString = 'OD'); //13453
    Ferme(QJal) ;
(*  if TEcr<>Nil then
  begin
    VideListe(Tecr);
    TEcr.Free ;
  end;*)
  TEcr := TList.Create ;
  TGen := TList.Create ;
  TAux := TList.Create ;
  ConstitueListesTVA(M,TEcr,TGen,TAux,Nbl) ;
  Etat := EtudieTVA(TEcr,TGen,TAux,Achat, OD);
  TEcr.Free ;
  TGen.Free ;
  TAux.Free ;
  Result := Etat ;
end;

procedure TOF_CPCONTRTVA.MajEtat(M: RMVT; EtatTVA: String3; NbL: integer);
var SQL : String ;
begin
  Try
    BEGINTRANS ;
    SQL := 'UPDATE ECRITURE SET E_CONTROLETVA="' + EtatTVA + '", E_DATEMODIF="'+UsTime(NowFutur)+'" ' + 'Where ' + WhereEcriture(tsGene,M,False) ;
    ExecuteSQL(SQL);
    COMMITTRANS;
  except
    V_PGI.IOError := oeUnknown ;
    ROLLBACK;
  end;
end;

function TOF_CPCONTRTVA.PasDeCodeTVA(TGen, TEcr: TList): boolean;
var
  i    : integer ;
  CGen : TGGeneral ;
  OEcr : TOBM ;
begin
  PasDeCodeTVA := True;
  for i:=0 to TGen.Count-1 do begin
    CGen := TGGeneral(TGen[i]) ;
//    if ((TGGeneral(TGen[i]).General<>VH^.EccEuroDebit) and (TGGeneral(TGen[i]).General<>VH^.EccEuroCredit)) then
      if NatureHTTVA(CGen.NatureGene) then begin
        OEcr := TOBM(TEcr[i]) ;
        if OEcr.GetMvt('E_TVA') = '' then Exit ;
      end;
  end;
  PasDeCodeTVA := False ;
end;

function TOF_CPCONTRTVA.EtudieTVA(TEcr, TGen, TAux: TList; Achat, OD: boolean): String3;
var
  RegTVA,CodeTva,CodeTpf : String3 ;
  SoumisTPF,CoherTVA,EncON : boolean ;
  i                  : integer ;
  CGen               : TGGeneral ;
  XX                 : X_TVA ;
  OHT                : TOBM ;
  Diff,STaxeA,STaxeD,Tolere : Double ;
  tsm                : TSorteMontant ;
  Ladec    : integer ;
  O    : TOBM ;
  szNature : String;
  szProrata          : String;
  XXTauxTva : Double ;
begin
  Result := 'COR' ;
  CoherTVA:=True ;
  Tolere := Valeur(GetControlText('TOLERECALC')) ;

  // Tests premier niveau
  if PasDeTiers(TGen,Taux)      then begin Result := 'STI'; Exit; end;
  if MultiTiers(TEcr,TGen,Taux) then begin Result := 'MTI'; Exit; end;
  if PasDeHT(TGen)              then begin Result := 'SHT'; Exit; end;
  if PasDeCodeTVA(TGen,TEcr)    then begin Result := 'STV'; Exit; end;
  if TiersSansRegime(TGen,TAux) then begin Result := 'RTI'; Exit; end;

  // Cohérence de la TVA
  RecupInfosTVA(TEcr,TGen,TAux,RegTVA,SoumisTPF) ;
  if OD then begin // 13453
    O := TOBM(TEcr[0]);
    if O<>Nil then begin
      szNature := O.GetMvt('E_NATUREPIECE');
      if (szNature='AC') or (szNature='FC') then Achat := False
                                            else Achat := True;
    end;
  end;
  for i:=0 to TGen.Count-1 do begin
    CGen := TGGeneral(TGen[i]) ;
//    if ((CGen.General<>VH^.EccEuroDebit) and (CGen.General<>VH^.EccEuroCredit)) then
      if NatureHTTVA(CGen.NatureGene) then begin
        OHT := TOBM(TEcr[i]) ;
        CodeTva := OHT.GetMvt('E_TVA') ;
        CodeTpf := OHT.GetMvt('E_TPF') ;
        EncON := (OHT.GetMvt('E_TVAENCAISSEMENT')='X') ;
        RenseigneHTByTva(CGen,RegTVA,CodeTva,CodeTpf,SoumisTPF,Achat,EncON,CoherTva) ;
      end;
      if Not CoherTVA then begin Result := 'INC'; Exit; end;
  end;

  // Calculs
  RempliTVA(TEcr,TGen,RegTVA,Achat,SoumisTPF) ;
  for i:=0 to TT_TVA.Count-1 do begin
    XX := X_TVA(TT_TVA[i]) ;

    // Tests de cohérence sur monnaie Pivot
    if ((XX.CHT<>'') and (XX.XHTP<>0) and (XX.XTVAAP<>0) and (XX.CTVA='')) then begin Result := 'STV'; Exit; end;
    if ((XX.CHT<>'') and (XX.XHTP<>0) and (XX.XTPFAP<>0) and (XX.CTPF='')) then begin Result := 'STV'; Exit; end;
    if ((XX.XTVAAP<>0) and (XX.CTVA<>'')) then if Not PresentTVA(XX.CTVA,TEcr) then begin Result := 'ERR'; Exit; end;
    if ((XX.XTPFAP<>0) and (XX.CTPF<>'')) then if Not PresentTVA(XX.CTPF,TEcr) then begin Result := 'ERR'; Exit; end;

    // GP : Reprise 18358 : vu avec CAyel : correction aberrante.
    {FQ18358 YMO 20/06/2006 Il n'y a plus un X_TVA par compte géné, mais un compte par ligne (on n'additionne plus
    les lignes d'un même compte, voir modif dans RempliTVA) ; Pour vérifier si on n'a pas un compte exonéré,
    il faut dorénavant vérifier tous les X_TVA concernant ce compte (toutes les lignes)   }
    (*
    TotalTVA:=0;
    for k := 0 to TT_TVA.Count-1 do begin
        XY := X_TVA(TT_TVA[k]);
        if (XY.CHT=XX.CHT) {on reprend celui en cours, c normal}
        and (XY.CHT<>'') and (XY.XHTP<>0) and (XY.CodeTva<>'') then
          TotalTVA:=TotalTVA + Tva2Taux(XY.RegTVA,XY.CodeTva,XY.Achat);
    end;
    if TotalTVA=0 then begin Result := 'EXO'; Exit; end;
    *)
     XXTauxTva:=Tva2Taux(XX.RegTVA,XX.CodeTva,XX.Achat) ; If Arrondi(XXTauxTva,3)=0 Then Continue ;
    // Tests de montant sur monnaie de saisie
    if DEV.Code<>V_PGI.DevisePivot then begin LaDec := DEV.Decimale; tsm := tsmDevise; end else
        begin LaDec := V_PGI.OkDecV ; tsm := tsmPivot ; end ;
    STaxeA := SommeCalc(XX.CTVA,True,True,tsm) ;
    STaxeD := SommeCalc(XX.CTVA,True,False,tsm) ;
    Diff := Abs(STaxeA-STaxeD) ;
    if Arrondi(Diff,LaDec)>Diff then Diff := Arrondi(Diff,LaDec) ;

    if Abs(Arrondi(SommeCalc(XX.CTVA,True,False,tsm)-SommeSais(XX.CTVA,TEcr,tsm),LaDec))>Diff+Tolere then begin
      szProrata := TOBM(TEcr[i]).GetMvt('E_QUALIFORIGINE');
      if (szProrata = 'TV1') then begin Result:='PR1'; Exit; end else
      if (szProrata = 'TVI') then begin Result:='PR2'; Exit; end else
      if (szProrata = 'TV2') then begin Result:='PR3'; Exit; end
                             else begin Result:='MTV'; Exit; end ;
    end;

    STaxeA := SommeCalc(XX.CTPF,False,True,tsm) ;
    STaxeD := SommeCalc(XX.CTPF,False,False,tsm) ;
    Diff := Abs(STaxeA-STaxeD) ;
    if Arrondi(Diff,LaDec)>Diff then Diff := Arrondi(Diff,LaDec) ;
    if Abs(Arrondi(SommeCalc(XX.CTPF,False,False,tsm)-SommeSais(XX.CTPF,TEcr,tsm),LaDec))>Diff+Tolere then begin Result:='MTV' ; Exit ; end ;
  end;
end;

function TOF_CPCONTRTVA.PasDeTiers(TGen, TAux: TList): boolean;
var
  i : integer;
begin
  PasDeTiers := False;
  for i := 0 to TGen.Count-1 do
    if EstTiersGeneTVA(TGGeneral(TGen[i])) then Exit;
  for i := 0 to TAux.Count-1 do
    if TGTiers(TAux[i])<>Nil then Exit;
  PasDeTiers := True;
end;

function TOF_CPCONTRTVA.PasDeHT(TGen: TList): boolean;
var
  i : integer;
begin
  PasDeHT := False;
  for i := 0 to TGen.Count-1 do
//    if ((TGGeneral(TGen[i]).General<>VH^.EccEuroDebit) and (TGGeneral(TGen[i]).General<>VH^.EccEuroCredit)) then
      if NatureHTTVA(TGGeneral(TGen[i]).NatureGene) then Exit;
  PasDeHT := True;
end;

function TOF_CPCONTRTVA.MultiTiers(TEcr, TGen, TAux: TList): boolean;
var
  i,Nb : integer;
begin
  Nb := 0;
  for i := 0 to TEcr.Count-1 do if TOBM(TEcr[i]).GetMvt('E_NUMECHE')<=1 then begin
    if EstTiersGeneTVA(TGGeneral(TGen[i])) then Inc(Nb);
    if TGTiers(TAux[i])<>Nil then Inc(Nb);
  end;
  MultiTiers := (Nb>1);
end;

function TOF_CPCONTRTVA.TiersSansRegime(TGen, TAux: TList): boolean;
var
  CGen : TGGeneral;
  CAux : TGTiers;
  i    : integer;
begin
  TiersSansRegime := True;
  for i := 0 to TGen.Count-1 do begin
    CGen := TGGeneral(TGen[i]);
    if EstTiersGeneTVA(CGen) then if CGen.RegimeTVA='' then Exit;
  end;

  for i := 0 to TAux.Count-1 do begin
    CAux := TGTiers(TAux[i]);
    if CAux<>Nil then if CAux.RegimeTVA='' then Exit;
  end;
  TiersSansRegime := False;
end;

procedure TOF_CPCONTRTVA.RempliTVA(TEcr, TGen: TList; RegTVA: String3; Achat, SoumisTPF: boolean);
var
  O   : TOBM;
  i   : integer;
  XX  : X_TVA;
  CGen : TGGeneral;
begin
  VideListe(TT_TVA);
  for i := 0 to TEcr.Count-1 do begin
    CGen := TGGeneral(TGen[i]);
    O := TOBM(TEcr[i]);
//    if ((TGGeneral(TGen[i]).General<>VH^.EccEuroDebit) and (TGGeneral(TGen[i]).General<>VH^.EccEuroCredit)) then
      if NatureHTTVA(CGen.NatureGene) then begin
      {YMO FQ18358 Dans le cas d'un même compte général présent sur 2 lignes différentes dans la pièce,
      on ajoute ici les montants sur le même XX, une seule entrée dans TT_TVA, avec une seule info
      concernantr la TVA. Or, on peut forcer 2 regimes de TVA différents sur ces deux lignes (de même compte).
      Un seul enreg ne suffit plus au moment où l'on calcule les montants de TVA pour la vérif dans la procédure
      EtudieTVA, qui s'emmêle alors les pinceaux...
      Le + simple est de laisser la création de deux X_TVA pour un même compte.

        Trouv := False;
        for k := 0 to TT_TVA.Count-1 do begin
          XX := X_TVA(TT_TVA[k]);
          if XX.CHT=CGen.General then begin Trouv := True; CumulTVA(XX,O,CGen,RegTVA,Achat,SoumisTPF); Break; end;
        end;
        if not Trouv then }
        begin
          XX := X_TVA.Create;
          XX.RegTVA := RegTVA;
          XX.Achat := Achat;
          XX.CodeTva := O.GetMvt('E_TVA');
          XX.XHTP := 0; XX.XTVADP := 0; XX.XTVAAP := 0; XX.XTPFDP := 0; XX.XTPFAP := 0;
//          XX.XHTE := 0; XX.XTVADE := 0; XX.XTVAAE := 0; XX.XTPFDE := 0; XX.XTPFAE := 0;
          XX.XHTD := 0; XX.XTVADD := 0; XX.XTVAAD := 0; XX.XTPFDD := 0; XX.XTPFAD := 0;
          CumulTVA(XX,O,CGen,RegTVA,Achat,SoumisTPF);
          TT_TVA.Add(XX);
        end;
      end;
  end;
end;

procedure TOF_CPCONTRTVA.CumulTVA(XX: X_TVA; O: TOBM; CGen: TGGeneral; RegTVA: String3; Achat, SoumisTPF: boolean);
var
  CodeTva,CodeTpf : String;
begin
XX.CHT := O.GetMvt('E_GENERAL');
// HT
XX.XHTP := XX.XHTP+O.GetMvt('E_DEBIT') -     O.GetMvt('E_CREDIT');
XX.XHTD := XX.XHTD+O.GetMvt('E_DEBITDEV') -  O.GetMvt('E_CREDITDEV');
//XX.XHTE := XX.XHTE+O.GetMvt('E_DEBITEURO') - O.GetMvt('E_CREDITEURO');
// TVA
CodeTva := O.GetMvt('E_TVA');
CodeTpf := O.GetMvt('E_TPF');
XX.CTVA := CGen.CpteTVA;
XX.XTVADP := HT2TVA(XX.XHTP,RegTVA,SoumisTPF,CodeTVA,CodeTPF,Achat,5); XX.XTVAAP := Arrondi(XX.XTVADP,V_PGI.OkDecV);
//XX.XTVADE := HT2TVA(XX.XHTE,RegTVA,SoumisTPF,CodeTVA,CodeTPF,Achat,5); XX.XTVAAE := Arrondi(XX.XTVADE,V_PGI.OkDecE);
XX.XTVADD := HT2TVA(XX.XHTD,RegTVA,SoumisTPF,CodeTVA,CodeTPF,Achat,5); XX.XTVAAD := Arrondi(XX.XTVADD,DEV.Decimale);
// TPF
XX.CTPF := CGen.CpteTPF;
XX.XTPFDP := HT2TPF(XX.XHTP,RegTVA,SoumisTPF,CodeTVA,CodeTPF,Achat,5); XX.XTPFAP := Arrondi(XX.XTPFDP,V_PGI.OkDecV);
//XX.XTPFDE := HT2TPF(XX.XHTE,RegTVA,SoumisTPF,CodeTVA,CodeTPF,Achat,5); XX.XTPFAE := Arrondi(XX.XTPFDE,V_PGI.OkDecE);
XX.XTPFDD := HT2TPF(XX.XHTD,RegTVA,SoumisTPF,CodeTVA,CodeTPF,Achat,5); XX.XTPFAD := Arrondi(XX.XTPFDD,DEV.Decimale);
end;

function TOF_CPCONTRTVA.SommeSais(CPT: String; TEcr: TList; tsm: TSorteMontant): Double;
var
  d : double;
  O : TOBM;
  i  : integer;
begin
  D := 0;
  for i := 0 to TEcr.Count-1 do begin
    O := TOBM(TEcr[i]);
    if O.GetMvt('E_GENERAL') = CPT then begin
      Case tsm of
        tsmDevise : D := D + O.GetMvt('E_DEBITDEV')  - O.GetMvt('E_CREDITDEV');
        tsmPivot  : D := D + O.GetMvt('E_DEBIT')     - O.GetMvt('E_CREDIT');
//        tsmEuro   : D := D + O.GetMvt('E_DEBITEURO') - O.GetMvt('E_CREDITEURO');
      end;
    end;
  end;
  Result := D;
end;

function TOF_CPCONTRTVA.SommeCalc(CPT: String; OkTVA, Arr: Boolean; tsm: TSorteMontant): Double;
var
  d : double;
  XX : X_TVA;
  i  : integer;
begin
  D := 0;
  for i := 0 to TT_TVA.Count-1 do begin
    XX := X_TVA(TT_TVA[i]);
    if ((OkTVa) and (XX.CTVA=CPT)) then begin
      Case tsm of
        tsmDevise : if Arr then D := D + XX.XTVAAD else D := D + XX.XTVADD;
        tsmPivot  : if Arr then D := D + XX.XTVAAP else D := D + XX.XTVADP;
//        tsmEuro   : if Arr then D := D + XX.XTVAAE else D := D + XX.XTVADE;
      end;
    end;
    if ((not OkTVa) and (XX.CTPF=CPT)) then begin
      Case tsm of
        tsmDevise : if Arr then D := D + XX.XTPFAD else D := D + XX.XTPFDD;
        tsmPivot  : if Arr then D := D + XX.XTPFAP else D := D + XX.XTPFDP;
//        tsmEuro   : if Arr then D := D + XX.XTPFAE else D := D + XX.XTPFDE;
      end;
    end;
  end;
  Result := D;
end;

function TOF_CPCONTRTVA.PresentTVA(CPT: String; TEcr: TList): boolean;
var
  i : integer;
begin
  Result := True;
  for i := 0 to TEcr.Count-1 do
    if TOBM(TEcr[i]).GetMvt('E_GENERAL') = CPT then Exit;
  Result := False;
end;

procedure TOF_CPCONTRTVA.BOuvrirClick(Sender: TObject);
var
  Nb : longint;
  ii : integer;
  //io : TIOErr;
  Timer, timer2 : TDateTime ;
  str : string ;
begin
  Timer := Now ;
  if not FListe.AllSelected then begin
    Nb := FListe.NbSelected;
    if Nb<=0 then
    begin HM.Execute(1,'','');
      Exit;
    end; // Vous n''avez sélectionné aucune pièce.
  end;
  V_PGI.IOError := oeUnknown ;
  // Confirmez-vous le traitement des pièces sélectionnées ?
  ii := HM.Execute(0,'','');
  Application.ProcessMessages;
  FreeAndNil(TT_TVA);
  Case ii of
    mrYes : begin
            TT_TVA := TList.Create;
            NowFutur := NowH;
            if V_PGI.SAV then Timer := Time ;
            //io := Transactions(ControleLaTva,3);
            Transactions(ControleLaTva,3);
            if V_PGI.SAV then
            begin
              timer2 := Time ;
              Str := 'De ' + TimeToStr(Timer) + ' à ' + TimeToStr(timer2) + #10#13 + 'Durée = ' + TimeToStr(Timer2- Timer) ;
              MESSAGEALERTE(str);
            end ;
            //if io<>oeOk then MessageAlerte(HM.Mess[2]); // Certaines pièces en cours de traitement ont stoppé la validation.
          end;
    mrNo  : FListe.ClearSelected;
    mrCancel :;
  end;
  if TT_TVA<>Nil then VideListe(TT_TVA);
  FreeAndNil(TT_TVA);
  if FListe.AllSelected then
  begin
    if HM.Execute(3,'','') = mryes  then
      BChercheClick(Nil)
    else
      FListe.ClearSelected;
  end
  else
  begin
    PGIInfo('Traitement terminé.');
    //BChercheClick(Nil) ;
  end;
end;

procedure TOF_CPCONTRTVA.InitMsgBox;
begin
HM.Mess.Add('0;'+Ecran.Caption+';Confirmez-vous le traitement des pièces sélectionnées ?;Q;YNC;Y;Y;');
HM.Mess.Add('1;'+Ecran.Caption+';Vous n''avez sélectionné aucune pièce.;E;O;O;O;');
HM.Mess.Add('Certaines pièces en cours de traitement ont stoppé la validation.');
HM.Mess.Add('3;'+Ecran.Caption+';Traitement terminé. Voulez vous rafraîchir la liste ?;Q;YN;Y;Y;');
end;

procedure TOF_CPCONTRTVA.BChercheClick(Sender: TObject);
begin
  TVAOKOnClick(nil);
  TFMul(Ecran).BChercheClick(Sender);
(*{$IFDEF EAGLCLIENT}
  GereSelectionsGrid(FListe,TFMul(Ecran).Q.TQ);
{$ELSE}
  GereSelectionsGrid(FListe,TFMul(Ecran).Q);
{$ENDIF}
*)
end;

procedure TOF_CPCONTRTVA.FListeDblClick(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
  TFMul(Ecran).Q.TQ.Seek(FListe.Row-1);
  TrouveEtLanceSaisie(TFMul(Ecran).Q.TQ,taModif,GetControlText('E_QUALIFPIECE'));
{$ELSE}
  TrouveEtLanceSaisie(TFMul(Ecran).Q,taModif,GetControlText('E_QUALIFPIECE'));
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 12/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONTRTVA.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;

procedure TOF_CPCONTRTVA.BParamTPFClick(Sender: TObject);
begin
  ParamTVATPF(False);
end;

procedure TOF_CPCONTRTVA.BParamTVAClick(Sender: TObject);
begin
  ParamTVATPF(True);
end;

procedure TOF_CPCONTRTVA.BZoomPieceClick(Sender: TObject);
begin
  FListeDblClick(Nil);
end;

procedure TOF_CPCONTRTVA.BSelectAllClick(Sender: TObject);
//var
//  Fiche : TFMul;
begin
  //Fiche := TFMul(Ecran);
  {$IFDEF EAGLCLIENT}
  if TFMul(Ecran).bSelectAll.Down then
    if not TFMul(Ecran).FetchLesTous then
    begin
      TFMul(Ecran).bSelectAllClick(nil);
      TFMul(Ecran).bSelectAll.Down := False;
      Exit;
    end;
  {$ELSE}
  // A FAIRE
  {$ENDIF}
  TFMul(Ecran).bSelectAllClick(nil);
end;

procedure TOF_CPCONTRTVA.E_ETABLISSEMENTEnter(Sender: TObject);
begin
  CSelectionTextControl( Sender ) ;
end;

procedure TOF_CPCONTRTVA.TVAOKOnClick(Sender: TObject);
begin
  if TVAOK.Checked then SetControlText('XX_WHERETVAOK','E_CONTROLETVA="COR"') ;
  If TVAOK.Checked = false then
    if TVAOK.State= cbGrayed then SetControlText('XX_WHERETVAOK','')
    else SetControlText('XX_WHERETVAOK','E_CONTROLETVA<>"COR"') ;
end;


Initialization
  registerclasses ( [ TOF_CPCONTRTVA ] ) ;
end.
