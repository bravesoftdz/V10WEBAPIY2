{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 29/04/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPMULFACTEFF ()
Mots clefs ... : TOF;CPMULFACTEFF
*****************************************************************}
Unit CPMULFACTEFF_TOF ;

Interface

Uses StdCtrls, 
     Controls,
     Classes,
     Windows,           // VK_F10
{$IFDEF EAGLCLIENT}
     MaineAGL,          // AGLLanceFiche
     eMul,              // TFMul
{$ELSE}
     HDB,               // THDBGrid
     db,
     dbtables,
     FE_Main,           // AGLLanceFiche
     Mul,               // TFMul
{$ENDIF}
     Saisie,            // TrouveEtLanceSaisie
     forms,
     sysutils,
     ComCtrls,
     HCtrls,      // THGrid
     HEnt1,
     HMsgBox,
     SaisUtil,    // tModeSaisieEff, tModeSR, OnEffet, OnChq, OnCB, OnBqe, srFou
     UTOF,
     AGLInit,     // TheData, TheTob
     HQry,        // THQuery
     Ent1,        // ExoToDates, AfficheLeSolde
     HTB97,       // TToolBarButton97
     SaisBor,     // LanceSaisieFolio
     ED_Tools,    // VideListe
     UTOB,        // TOB
     Paramsoc;

Function PointeMvtCpt(Gene,Auxi,LibCpt : String ; LOBM : TList ; DateEcheDefaut : tDateTime ; NewFlux{,ModeOppose} : Boolean ; ModeSaisie : tModeSaisieEff ; ModeSR : tModeSR ; O : TOB) : Boolean ;

Type tTotSel = Record
              XDebitP,XDebitE,XCreditP,XCreditE : Double ;
              NbP : Integer ;
              End ;

Type
  TOF_CPMULFACTEFF = Class (TOF)
    HM : THMsgBox;
{$IFDEF EAGLCLIENT}
    FListe : THGrid;
{$ELSE}
    FListe : THDBGrid;
{$ENDIF}
    Q : THQuery;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure InitMsgBox;
  private
    Gene,Auxi,LibCpt : String ;
    LOBM : TList;
    DateEcheDefaut : TDateTime ;
    NewFlux : Boolean ;
    {ModeOppose : Boolean ;}
    ModeSaisie : tModeSaisieEff ;
    ModeSR : tModeSR ;
    TOB2 : TOB;

    LOBM1      : TList;
    TotSelInit : tTotSel;
    TotSel     : tTotSel;

    procedure E_EXERCICEChange(Sender: TObject);
    procedure bVoirTotSelClick(Sender: TObject);
    procedure bSelectAllClick(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure FListeFlipSelection(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BChercheClick(Sender: TObject);
    procedure AuxiElipsisClick(Sender : TObject);

    procedure InitCriteres ;
    procedure AfficheTotauxBasEcran ;
    Function  CtrlMontantSel(LOBM : TList) : Boolean ;
    procedure FAitOBMDrag ;
    procedure RefreshLOBM ;
    Function  ExistePasDansLOBM(O1 : TOBM ; TOBL1 : TOB) : Boolean ;
    procedure TraiteMajFliste2;
    procedure CalculMontantSelect ;
    function  ExisteDansTOB(O1 : TOBM ; TOBL1 : TOB) : Boolean;
    procedure ReinitTotSel;
    Procedure DeleteFromLOBM ;
    procedure InitLOBM1;
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
  , UTofMulParamGen; {23/04/07 YMO F5 sur Auxiliaire}

var
  InitEnCours : boolean;

Function PointeMvtCpt(Gene,Auxi,LibCpt : String ; LOBM : TList ; DateEcheDefaut : tDateTime ; NewFlux{,ModeOppose} : Boolean ; ModeSaisie : tModeSaisieEff ; ModeSR : tModeSR ; O : TOB) : Boolean ;
var
  szArg : String;
  szRetour : String;
begin
  szArg := Gene+';'+Auxi+';'+DateToStr(DateEcheDefaut);
  if NewFlux then szArg := szArg +';X' else szArg := szArg +';-';
  {if ModeOppose then szArg := szArg +';X' else szArg := szArg +';-';}
  Case ModeSaisie Of
    OnEffet : szArg := szArg +';0';
    OnChq   : szArg := szArg +';1';
    OnCB    : szArg := szArg +';2';
    OnBqe   : szArg := szArg +';3';
  end;
  Case ModeSR Of
    srRien : szArg := szArg +';0;';
    srCli  : szArg := szArg +';1;';
    srFou  : szArg := szArg +';2;';
  end;
  szArg := szArg + LibCpt;
  TheData := LOBM;
  TheTOB := O;
  szRetour := AGLLanceFiche('CP','CPMULFACTEFF','','', szArg);
  Result := (szRetour ='X');
end;

procedure TOF_CPMULFACTEFF.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPMULFACTEFF.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPMULFACTEFF.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPMULFACTEFF.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPMULFACTEFF.OnArgument (S : String ) ;
begin
  Inherited ;

  // Création des contrôles
  HM := THMsgBox.create(FMenuG);
  InitMsgBox;
  Q := TFMul(Ecran).Q ;

  // Récupère les arguments
  Gene := ReadTokenSt(S);
  Auxi := ReadTokenSt(S);
  DateEcheDefaut := StrToDate(ReadTokenSt(S));
  NewFlux := (ReadTokenSt(S) = 'X');
  {ModeOppose := (ReadTokenSt(S) = 'X');}
  Case ReadTokenI(S) Of
    0 : begin ModeSaisie := OnEffet; TFMul(Ecran).FNomFiltre := 'CPFACTCLIEFFT';   Q.Liste:='CPMULFACTCLIEFFT'; end;
    1 : begin ModeSaisie := OnChq;   TFMul(Ecran).FNomFiltre := 'CPMULFACTCLICHQ'; Q.Liste:='CPMULFACTCLICHQ';  end;
    2 : begin ModeSaisie := OnCB;    TFMul(Ecran).FNomFiltre := 'CPMULFACTCLICB';  Q.Liste:='CPMULFACTCLICB';   end;
    3 : begin ModeSaisie := OnBqe;   TFMul(Ecran).FNomFiltre := 'CPMULFACTCLIBQE'; Q.Liste:='CPMULFACTCLIBQE';  end;
  end;
  Case ReadTokenI(S) Of
    0 : ModeSR := srRien;
    1 : ModeSR := srCli;
    2 : ModeSR := srFou;
  end;
  LibCpt := S;
  LOBM := TList(TheData);
  TheData := nil;
  Tob2 := LaTob;
  LaTob := nil;

  // Evénements des contrôles
  THValComboBox(GetControl('E_EXERCICE',True)).OnChange := E_EXERCICEChange;
  TToolBarButton97(GetControl('BVOIRTOTSEL',True)).OnClick := bVoirTotSelClick;
  TToolBarButton97(GetControl('BSELECTALL',True)).OnClick := bSelectAllClick;
  TToolBarButton97(GetControl('BOUVRIR',True)).OnClick := BOuvrirClick;
{$IFDEF EAGLCLIENT}
  FListe := THGrid(TFMul(Ecran).FListe);
  THGrid(GetControl('FLISTE',True)).OnDblClick := FListeDblClick;
  THGrid(GetControl('FLISTE',True)).OnFlipSelection := FListeFlipSelection;
{$ELSE}
  FListe := THDBGrid(TFMul(Ecran).FListe);
  THDBGrid(GetControl('FLISTE',True)).OnDblClick := FListeDblClick;
  THDBGrid(GetControl('FLISTE',True)).OnFlipSelection := FListeFlipSelection;
{$ENDIF}
  THNumEdit(GetControl('FDEBITPA', True)).Masks.PositiveMask := '#,##0.00';
  THNumEdit(GetControl('FCREDITPA', True)).Masks.PositiveMask := '#,##0.00';
  THNumEdit(GetControl('FSOLDEA', True)).Masks.PositiveMask := '#,##0.00';
  Ecran.OnKeyDown := FormKeyDown;
  TFMul(Ecran).BCherche.OnClick := BChercheClick;

  // FormCreate
  LOBM1 := TList.Create;

  // FormShow
  InitCriteres ;

  if LOBM=NIL then begin
    Q.Manuel := False;
    Q.UpdateCriteres ;
  end;
  Ecran.Caption := HM.Mess[1]+' '+LibCpt+' ('+Gene+'/'+Auxi+')' ;

  THNumEdit(GetControl('FDEBITPA', True)).Value  := 0;
  THNumEdit(GetControl('FCREDITPA', True)).Value := 0;
  THNumEdit(GetControl('FSOLDEA', True)).Value   := 0;

  CalculMontantSelect ;

  TotSelInit := TotSel;
  AfficheTotauxBasEcran;
  SetControlVisible('FPIED', True);

end ;

procedure TOF_CPMULFACTEFF.OnClose ;
begin
  Inherited ;
  // FormCloseQuery
  if (Ecran.ModalResult = mrCancel) And (LOBM<>NIL) And (LOBM.Count>0) And (Arrondi(TotSelInit.XDEBITP+TotSelInit.XCREDITP-(TotSel.XDEBITP+TotSelInit.XCREDITP),V_PGI.OkDecV)<>0) then begin
    // Confirmez-vous les mouvements sélectionnés ?
    if HM.Execute(9,Ecran.Caption,'')<>MrYes Then begin
      LastError := 1;
      Exit;
    end;
  end;

  HM.Free;

  // FormClose
  if (Ecran.ModalResult = mrOK) then TFMul(Ecran).Retour := 'X';
  VideListe(LOBM1) ;
  LOBM1.Clear ;
  LOBM1.Free ;
end ;

procedure TOF_CPMULFACTEFF.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPMULFACTEFF.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPMULFACTEFF.InitMsgBox;
begin
  HM.Mess.Add('0;?caption?;Validation impossible : Vous n''avez pas sélectionnné les comptes de l''écriture en cours de saisie.;W;O;O;O;');
  HM.Mess.Add('Liste des mouvements du compte à sélectionner');
  HM.Mess.Add('2;?caption?;Validation impossible : Le montant global est créditeur.;W;O;O;O;');
  HM.Mess.Add('3;?Caption?;Confirmez-vous l''affectation des échéances sélectionnées ?;Q;YN;Y;N;');
  HM.Mess.Add('4;?Caption?;Confirmez-vous la suppression des échéances sélectionnées du lot ?;Q;YN;Y;N;');
  HM.Mess.Add('Liste des mouvements sélectionnés :');
  HM.Mess.Add('Débit');
  HM.Mess.Add('Crédit');
  HM.Mess.Add('Solde');
  HM.Mess.Add('9;?Caption?;Confirmez-vous les mouvements sélectionnés ?;Q;YN;Y;N;');
  HM.Mess.Add('10;?Caption?;Un ou plusieurs mouvements n''ont pas été pris en compte car ils sont déja sélectionnés sur une autre ligne.;W;O;O;O;');
end;

procedure TOF_CPMULFACTEFF.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var
  Vide : boolean ;
begin
  Vide:=(Shift=[]) ;
  if Vide then begin
    Case Key of
      VK_F10 : begin
                 Key:=0 ;
                 If Fliste.NbSelected>0 Then TraiteMajFListe2;
                 BOuvrirClick(Nil) ; Exit ;
               end;
      VK_F9  : ReinitTotSel;
    end;
  end;
  TFMul(Ecran).FormKeyDown(Sender, Key, Shift);
end;

procedure TOF_CPMULFACTEFF.E_EXERCICEChange(Sender: TObject);
begin
  if THValComboBox(GetControl('E_EXERCICE', True)).ItemIndex>0 then ExoToDates(GetControlText('E_EXERCICE'), GetControl('E_DATECOMPTABLE',True), GetControl('E_DATECOMPTABLE_',True) );
end;

procedure TOF_CPMULFACTEFF.bVoirTotSelClick(Sender: TObject);
begin
  SetControlVisible('FPIED', TToolBarButton97(GetControl('BVOIRTOTSEL', True)).Down);
end;

procedure TOF_CPMULFACTEFF.bSelectAllClick(Sender: TObject);
{$IFDEF EAGLCLIENT}
var
  Fiche : TFMul;
  i : Integer;
{$ENDIF}
begin
TFMul(Ecran).bSelectAllClick(nil);
{$IFDEF EAGLCLIENT}
Fiche := TFMul(Ecran);
if Fiche.bSelectAll.Down then
begin
  Fiche.FetchLesTous;
  ReinitTotSel;
  for i:=1 to FListe.RowCount-1 do begin
    FListe.Row := i;
    FListe.FlipSelection(i);
  end;
end
else
  ReinitTotSel;
//AfficheTotauxBasEcran;
{$ELSE}
  Q.First ;
  While Not Q.Eof Do BEGIN
    FListe.FlipSelection ;
    Q.Next ;
  END
{$ENDIF}
end;

procedure TOF_CPMULFACTEFF.BOuvrirClick(Sender: TObject);
begin
//  TFMul(Ecran).BOuvrirClick(Sender); // inherited;

  if Sender<>NIL then begin
    If (FListe.nbSelected > 0) then TraiteMajFListe2;
  end;

  if LOBM=NIL then Exit ;

  if (Gene <> GetControlText('E_GENERAL')) Or (Auxi <> GetControlText('E_AUXILIAIRE')) then begin
    // Validation impossible : Vous n''avez pas sélectionnné les comptes de l''écriture en cours de saisie.
    HM.execute(0,Ecran.Caption,'');
    Exit;
  end;

  if Not CtrlMontantSel(LOBM) then begin
    // Validation impossible : Le montant global est créditeur.
    HM.Execute(2,Ecran.Caption,'');
    Exit;
  end;

  Ecran.ModalResult := mrOk;
end;

procedure TOF_CPMULFACTEFF.FListeDblClick(Sender: TObject);
var
  sMode : String ;
begin
  TFMul(Ecran).FListeDblClick(Sender); // inherited;

{$IFDEF EAGLCLIENT}
  if ((Q.TQ.EOF) and (Q.TQ.BOF)) then Exit ;
{$ELSE}
  if ((Q.EOF) and (Q.BOF)) then Exit ;
{$ENDIF}

  sMode:=Q.FindField('E_MODESAISIE').AsString ;

{$IFDEF EAGLCLIENT}
  if ((sMode<>'') and (sMode<>'-')) then LanceSaisieFolio(Q.TQ, TFMul(Ecran).TypeAction)
                                    else TrouveEtLanceSaisie(Q.TQ,taConsult,'N') ;
{$ELSE}
  if ((sMode<>'') and (sMode<>'-')) then LanceSaisieFolio(Q, TFMul(Ecran).TypeAction)
                                    else TrouveEtLanceSaisie(Q,taConsult,'N') ;
{$ENDIF}
end;

procedure TOF_CPMULFACTEFF.FListeFlipSelection(Sender: TObject);
var
  Coeff : Integer ;
  D,C : Double ;
begin
  inherited;
  {$IFDEF EAGLCLIENT}
  if (FListe.RowCount <=1) or (InitEnCours) then Exit;
  {$ELSE}
  if (not FListe.DataSource.DataSet.EOF) or (InitEnCours) then Exit;
  {$ENDIF EAGLCLIENT}

  AfficheTotauxBasEcran ;
{$IFDEF EAGLCLIENT}
  if (Q.TQ = nil) or (Q.TQ.Detail.Count = 0) then Exit;
  If FListe.IsSelected(FListe.Row) then Coeff:=1 else Coeff:=-1 ;

  Q.TQ.Seek(FListe.Row-1);

  //YMO deleteLOBM
  If Not FListe.IsSelected(FListe.Row) Then
  BEGIN  // on supprime de LOBM
      DeleteFromLOBM ;
  END ;

{$ELSE}
  if (Q = nil) then Exit;
  If FListe.IsCurrentSelected then Coeff:=1 else Coeff:=-1 ;
{$ENDIF}

  if (TotSel.NbP=0) and (Coeff = -1) then Exit;
  TotSel.NbP := TotSel.NbP+Coeff;
  D := Q.FindField('E_DEBIT').AsFloat ;
  C := Q.FindField('E_CREDIT').AsFloat ;
  if (Arrondi(D,V_PGI.OkDecV) = 0) then C := C-Q.FindField('E_COUVERTURE').AsFloat
                                   else D := D-Q.FindField('E_COUVERTURE').AsFloat;
  TotSel.XDebitP  := Arrondi(TotSel.XDebitP+(D*Coeff),V_PGI.OkDecV) ;
  TotSel.XCreditP := Arrondi(TotSel.XCreditP+(C*Coeff),V_PGI.OkDecV) ;

{  If Q.FindField('E_DEBITEURO')<>NIL then begin
    D := Q.FindField('E_DEBITEURO').AsFloat ;
    If Arrondi(D,V_PGI.OkDecV)<>0 then D := D-Q.FindField('E_COUVERTUREEURO').AsFloat ;
    TotSel.XDebitE := Arrondi(TotSel.XDebitE+(D*Coeff),V_PGI.OkDecE) ;
  end;

  If Q.FindField('E_CREDITEURO')<>NIL Then begin
    C := Q.FindField('E_CREDITEURO').AsFloat;
    If Arrondi(C,V_PGI.OkDecV)<>0 Then C := C-Q.FindField('E_COUVERTUREEURO').AsFloat;
    TotSel.XCreditE := Arrondi(TotSel.XCreditE+(C*Coeff),V_PGI.OkDecE);
  end;}

  AfficheTotauxBasEcran;
end;


procedure TOF_CPMULFACTEFF.AfficheTotauxBasEcran;
var
  Debit, Credit : THNumEdit;
begin
  Debit  := THNumEdit(GetControl('FDEBITPA', True));
  Credit := THNumEdit(GetControl('FCREDITPA', True));
  // Totaux Pointés sur la référence en bas de l'écran
  {if ModeOppose then begin
    Debit.Value := TotSel.XDebitE;
    Credit.Value := TotSel.XCreditE;
    end
  else begin}
    Debit.Value := TotSel.XDebitP;
    Credit.Value := TotSel.XCreditP;
//  end;

  AfficheLeSolde(THNumEdit(GetControl('FSOLDEA', True)), Debit.Value , Credit.Value );
  SetControlText('FNBP', FloatToStr(TotSel.NbP));
end;

procedure TOF_CPMULFACTEFF.CalculMontantSelect;
var
  i : Integer ;
  O : TOBM ;
  D,C : Double ;
begin
  Fillchar(TotSel,SizeOf(TotSel),#0) ;
  for i:=0 To LOBM.Count-1 Do begin
    O := TOBM(LOBM[i]);
    TotSel.NbP := TotSel.NbP+1;
    D := O.GetMvt('E_DEBIT');
    C := O.GetMvt('E_CREDIT');

    If Arrondi(D,V_PGI.OkDecV) = 0 then C := C - O.GetMvt('E_COUVERTURE')
                                   else D := D - O.GetMvt('E_COUVERTURE');

    TotSel.XDebitP  := Arrondi(TotSel.XDebitP + (D),V_PGI.OkDecV);
    TotSel.XCreditP := Arrondi(TotSel.XCreditP + (C),V_PGI.OkDecV);

//    D := O.GetMvt('E_DEBITEURO');
//    If Arrondi(D,V_PGI.OkDecV) <> 0 then D := D - O.GetMvt('E_COUVERTUREEURO');
//    TotSel.XDebitE := Arrondi(TotSel.XDebitE + (D),V_PGI.OkDecE);
//    C := O.GetMvt('E_CREDITEURO');
//    If Arrondi(C,V_PGI.OkDecV) <> 0 Then C := C - O.GetMvt('E_COUVERTUREEURO');
//    TotSel.XCreditE := Arrondi(TotSel.XCreditE + (C),V_PGI.OkDecE);
  END ;
end;

function TOF_CPMULFACTEFF.CtrlMontantSel(LOBM: TList): Boolean;
var
  O : TOBM ;
  D,C,Couv : Double ;
  i : integer ;
  M : Double ;
begin
  Result := False;
  D:=0 ; C:=0 ; Couv:=0 ;

  If LOBM=NIL Then Exit ;
  If LOBM.Count<=0 Then Exit ;

  for i:=0 To LOBM.Count-1 do begin
    O := TOBM(LOBM[i]) ;
    D := Arrondi(D + O.GetMvt('E_DEBIT'),V_PGI.OkDecV) ;
    C := Arrondi(C + O.GetMvt('E_CREDIT'),V_PGI.OkDecE) ;
    Couv  := Arrondi(Couv + O.GetMvt('E_COUVERTURE'),V_PGI.OkDecV) ;
  end;
  Result := True;

  M := Arrondi(D-C-Couv,V_PGI.OkDecV) ;

  If Modesr=srFou then begin
    if M>=0 then Result := False;
    end
  else begin
    if M<=0 then Result := False;
  end;
end;

function TOF_CPMULFACTEFF.ExistePasDansLOBM(O1: TOBM; TOBL1: TOB): Boolean;
var
  O : TOBM ;
  i : Integer ;
begin
  Result := True;

  For i:=0 To LOBM.Count-1 do begin
    O:=TOBM(LOBM[i]) ;
    If O1<>NIL then begin
      If (O.GetMvt('E_JOURNAL')=O1.GetMvt('E_JOURNAL')) And
         (O.GetMvt('E_EXERCICE')=O1.GetMvt('E_EXERCICE')) And
         (O.GetMvt('E_DATECOMPTABLE')=O1.GetMvt('E_DATECOMPTABLE')) And
         (O.GetMvt('E_NUMEROPIECE')=O1.GetMvt('E_NUMEROPIECE')) And
         (O.GetMvt('E_NUMLIGNE')=O1.GetMvt('E_NUMLIGNE')) And
         (O.GetMvt('E_NUMECHE')=O1.GetMvt('E_NUMECHE')) And
         (O.GetMvt('E_QUALIFPIECE')=O1.GetMvt('E_QUALIFPIECE')) Then begin
        Result := False;
        Exit;
      end;
      end
    else begin
      If (O.GetMvt('E_JOURNAL')=TOBL1.GetValue('E_JOURNAL')) And
         (O.GetMvt('E_EXERCICE')=TOBL1.GetValue('E_EXERCICE')) And
         (O.GetMvt('E_DATECOMPTABLE')=TOBL1.GetValue('E_DATECOMPTABLE')) And
         (O.GetMvt('E_NUMEROPIECE')=TOBL1.GetValue('E_NUMEROPIECE')) And
         (O.GetMvt('E_NUMLIGNE')=TOBL1.GetValue('E_NUMLIGNE')) And
         (O.GetMvt('E_NUMECHE')=TOBL1.GetValue('E_NUMECHE')) And
         (O.GetMvt('E_QUALIFPIECE')=TOBL1.GetValue('E_QUALIFPIECE')) Then begin
        Result := False;
        Exit;
      end;
    end;
  end;
end;

function TOF_CPMULFACTEFF.ExisteDansTOB(O1 : TOBM ; TOBL1 : TOB) : Boolean;
Var
  O : TOB;
  i : Integer;
begin
  Result := False;
  for i := 0 To TOB2.Detail.Count-1 do begin
    O := TOB2.Detail[i];
    if O1<>NIL then begin
      If (O.GetValue('E_JOURNAL')      = O1.GetMvt('E_JOURNAL')) And
         (O.GetValue('E_EXERCICE')     = O1.GetMvt('E_EXERCICE')) And
         (O.GetValue('E_DATECOMPTABLE')= O1.GetMvt('E_DATECOMPTABLE')) And
         (O.GetValue('E_NUMEROPIECE')  = O1.GetMvt('E_NUMEROPIECE')) And
         (O.GetValue('E_NUMLIGNE')     = O1.GetMvt('E_NUMLIGNE')) And
         (O.GetValue('E_NUMECHE')      = O1.GetMvt('E_NUMECHE')) And
         (O.GetValue('E_QUALIFPIECE')  = O1.GetMvt('E_QUALIFPIECE')) Then BEGIN
        Result := True;
        Exit ;
      end;
      end
    else begin
      If (O.GetValue('E_JOURNAL')       = TOBL1.GetValue('E_JOURNAL')) And
         (O.GetValue('E_EXERCICE')      = TOBL1.GetValue('E_EXERCICE')) And
         (O.GetValue('E_DATECOMPTABLE') = TOBL1.GetValue('E_DATECOMPTABLE')) And
         (O.GetValue('E_NUMEROPIECE')   = TOBL1.GetValue('E_NUMEROPIECE')) And
         (O.GetValue('E_NUMLIGNE')      = TOBL1.GetValue('E_NUMLIGNE')) And
         (O.GetValue('E_NUMECHE')       = TOBL1.GetValue('E_NUMECHE')) And
         (O.GetValue('E_QUALIFPIECE')   = TOBL1.GetValue('E_QUALIFPIECE')) Then BEGIN
        Result := True;
        Exit ;
      end;
    end;
  end;
end;

procedure TOF_CPMULFACTEFF.FAitOBMDrag;
var
  i : Integer ;
  Q1 : TQuery ;
  O : TOBM ;
begin
  If (Gene <> GetControlText('E_GENERAL')) Or (Auxi <> GetControlText('E_AUXILIAIRE')) then begin
    // Validation impossible : Vous n''avez pas sélectionnné les comptes de l''écriture en cours de saisie.
    HM.execute(0,Ecran.Caption,'') ;
    Exit ;
  end;

  VideListe(LOBM1) ;
  LOBM1.Clear ;

  for i:=0 to FListe.NbSelected-1 do
  begin
    FListe.GotoLeBookmark(i);
{$IFDEF EAGLCLIENT}
    Q.TQ.Seek(FListe.Row-1);
{$ENDIF}

    Q1:=OpenSQL('SELECT * FROM ECRITURE WHERE E_JOURNAL="'+Q.FindField('E_JOURNAL').AsString+'"'
               +' AND E_EXERCICE="'+QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime))+'"'
               +' AND E_DATECOMPTABLE="'+USDATETIME(Q.FindField('E_DATECOMPTABLE').AsDateTime)+'"'
               +' AND E_NUMEROPIECE='+Q.FindField('E_NUMEROPIECE').AsString
               +' AND E_NUMLIGNE='+Q.FindField('E_NUMLIGNE').AsString
               +' AND E_NUMECHE='+Q.FindField('E_NUMECHE').AsString
               +' AND E_GENERAL="'+Gene+'" '
               +' AND E_AUXILIAIRE="'+Auxi+'" '
               +' AND E_QUALIFPIECE="N" ',True) ;
    if Not Q1.Eof then begin
      O:=TOBM.Create(EcrGen,'',False) ;
      O.ChargeMvt(Q1) ;
      LOBM1.Add(O) ;

    end;
    Ferme(Q1) ;
  end;
end;

procedure TOF_CPMULFACTEFF.InitCriteres;
var
  dd : TDateTime ;
begin
  SetControlText('E_GENERAL', Gene);
  SetControlText('E_AUXILIAIRE', Auxi);
  SetControlProperty('E_EXERCICE', 'ITEMINDEX', 0);
  SetControlText('E_DEVISE', V_PGI.DevisePivot);

{b fb 30/11/2005 FQ17094}
  SetControlText('E_DATECOMPTABLE', DateToStr(DebutDeMois(V_PGI.DateEntree)));
  SetControlText('E_DATECOMPTABLE_', DateToStr(FinDeMois(V_PGI.DateEntree)));
{e fb 30/11/2005 FQ17094}
  Case ModeSaisie Of
    OnEffet : begin
              DD:=DebutDeMois(DateEcheDefaut); SetControlText('E_DATEECHEANCE', DateToStr(DD));
              DD:=FinDeMois(DateEcheDefaut)  ; SetControlText('E_DATEECHEANCE_', DateToStr(DD));
              end;
    OnBqe   : begin
              SetControlEnabled('E_DEVISE', True);
              SetControlEnabled('TE_DEVISE', True);
              end;
  end;

  Case ModeSR Of
    srFou : begin
            SetControlText('XX_WHERENTP', 'E_NATUREPIECE="FF" OR E_NATUREPIECE="AF" OR E_NATUREPIECE="OD"  OR E_NATUREPIECE="OF"');
            end;
  end;
end;

procedure TOF_CPMULFACTEFF.RefreshLOBM;
var
  i : Integer ;
  O1,O : TOBM ;
  bMessage : Boolean;
begin
  bMessage := True;
  for i:=0 To LOBM1.Count-1 do begin
    O1 := TOBM(LOBM1[i]);
    O := NIL ;
    if ExisteDansTOB(O1,nil) then begin if bMessage then HM.Execute(10,Ecran.Caption,''); bMessage := False; Continue; end;
    If ExistePasDansLOBM(O1,NIL) then begin
      EgaliseOBM(O1,O) ;
      LOBM.Add(O) ;
    end;
  end;
end;

procedure TOF_CPMULFACTEFF.TraiteMajFliste2;
begin
  FaitOBMDrag ;
  RefreshLOBM ;

  //FListe.ClearSelected ;
  CalculMontantSelect ;
  AfficheTotauxBasEcran ;

end;

procedure TOF_CPMULFACTEFF.ReinitTotSel;
begin
  FillChar(TotSel,Sizeof(TotSel),#0);
end;

procedure TOF_CPMULFACTEFF.BChercheClick(Sender: TObject);
begin
  TFMul(Ecran).bSelectAll.Down := False;
// YMO FQ 16532 déclenché à l'ouverture de la fiche,
// empêchait le suivi des mouvements sélectionnés
//  ReinitTotSel;
  TFMul(Ecran).BChercheClick(Sender);
  InitLOBM1;
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 23/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
Procedure TOF_CPMULFACTEFF.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;


Procedure TOF_CPMULFACTEFF.DeleteFromLOBM ;
Var O,OToKill : TOBM ;
    i,Ind : Integer ;
BEGIN

Ind:=-1 ;
For i:=0 To LOBM.Count-1 Do
  BEGIN
  O:=TOBM(LOBM[i]) ;
  
  If (O.GetMvt('E_JOURNAL')=Q.FindField('E_JOURNAL').AsString) And
     (O.GetMvt('E_EXERCICE')=QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime))) And
     (O.GetMvt('E_DATECOMPTABLE')=Q.FindField('E_DATECOMPTABLE').AsDateTime) And
     (O.GetMvt('E_NUMEROPIECE')=Q.FindField('E_NUMEROPIECE').AsString) And
     (O.GetMvt('E_NUMLIGNE')=Q.FindField('E_NUMLIGNE').AsString) And
     (O.GetMvt('E_NUMECHE')=Q.FindField('E_NUMECHE').AsString) And
     (O.GetMvt('E_QUALIFPIECE')='N')
     Then BEGIN Ind:=i ; Break ; END ;

  END ;
If Ind<0 Then Exit ;
OToKill:=TOBM(LOBM[Ind]) ;
OToKill.Free ;
LOBM.Delete(Ind) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Yann MORENO
Créé le ...... : 24/10/2005
Modifié le ... :   /  /
Description .. : Initialisation : on 'flip' dans la grid les lignes précedemment
Suite ........ : sélectionnées
Mots clefs ... :
*****************************************************************}
Procedure TOF_CPMULFACTEFF.InitLOBM1 ;
Var i : Integer ;
    OLst : TOBM ;
//    QListe : TQuery ;
begin

  for i:=0 to FListe.RowCount-2 do //la 1ere ligne du grid est prise en compte
  begin

{$IFDEF EAGLCLIENT}
 //   Q.TQ.Seek(i);
{$ENDIF}

 {   QListe:=OpenSQL('SELECT * FROM ECRITURE WHERE E_JOURNAL="'+Q.FindField('E_JOURNAL').AsString+'"'
               +' AND E_EXERCICE="'+QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime))+'"'
               +' AND E_DATECOMPTABLE="'+USDATETIME(Q.FindField('E_DATECOMPTABLE').AsDateTime)+'"'
               +' AND E_NUMEROPIECE='+Q.FindField('E_NUMEROPIECE').AsString
               +' AND E_NUMLIGNE='+Q.FindField('E_NUMLIGNE').AsString
               +' AND E_NUMECHE='+Q.FindField('E_NUMECHE').AsString
               +' AND E_GENERAL="'+Gene+'" '
               +' AND E_AUXILIAIRE="'+Auxi+'" '
               +' AND E_QUALIFPIECE="N" ',True) ;
    if Not QListe.Eof then }
    begin
      OLst:=TOBM.Create(EcrGen,'',False) ;
//      OLst.ChargeMvt(QListe) ;

      OLst.GetLigneGrid(Fliste,i+1,'E_JOURNAL;E_DATECOMPTABLE;E_NUMEROPIECE;;;;;;;E_NUMLIGNE;E_NUMECHE;') ;
      OLst.PutValue('E_EXERCICE',QuelExo(OLst.GetValue('E_DATECOMPTABLE'))) ;
      OLst.PutValue('E_QUALIFPIECE','N') ;

      // utilise la fonction existante (d'où le Not...pas)
      If Not ExistePasDansLOBM(OLst,NIL) then
      begin
//          FListe.Row := i+1;
          InitEnCours:=true;
          FListe.FlipSelection(i+1);
          InitEnCours:=False;
          LOBM1.Add(OLst);
      end;

    end;

  //  Ferme(QListe) ;
  end;

end;

Initialization
  registerclasses ( [ TOF_CPMULFACTEFF ] ) ;
end.



