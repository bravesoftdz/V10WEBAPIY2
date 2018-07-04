{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 05/05/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPLETREGUL ()
Mots clefs ... : TOF;CPLETREGUL
*****************************************************************}
Unit CPLETREGUL_TOF;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     CPJOURNAL_TOM,     // FicheJournal
{$IFDEF EAGLCLIENT}
     MaineAGL,          // AGLLanceFiche
     eMul,              // TFMul
{$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     HDB,               // THDBGrid
     FE_Main,           // AGLLanceFiche
     Mul,               // TFMul
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,      // TToolBarButton97
     Ent1,       // tProfilTraitement
     SAISUTIL,   // RDEVISE, RLETTR, RMVT, TOBM
     LettUtil,   // REGJAL, PG_Lettrage, GoReqMajLet
     HQry,       // THQuery
     UTOb,       // TQuery, Tob
     utilPGI,    // Resolution,
     HStatus,    // InitMove, MoveCur, FiniMove
     SaisComm,   // tsmPivot, tsmEuro
     DelVisuE,   // VisuPiecesGenere
     ed_tools,   // VideListe
     LetBatch,   // CreerPartieDoubleLett
     UTOF ;

procedure RegulLettrage( Regul,Convert : Boolean );
procedure RegulInverse;
procedure RegulLettrageMP( Regul,Convert : Boolean ; Qui : tProfilTraitement);
procedure RegulLettrageMPAuto( Regul,Convert : Boolean );
function StrToPT(psz : String) : tProfilTraitement;
function PTToStr(pPT : tProfilTraitement) : String;

Type
  TOF_CPLETREGUL = Class (TOF)
    HM : THMsgBox;
{$IFDEF EAGLCLIENT}
    GR : THGrid;
{$ELSE}
    GR : THDBGrid;
{$ENDIF}
    QR : THQuery;
    ModeOppose : TCheckBox;
    MAXDEBIT   : THNumEdit;
    MAXCREDIT  : THNumEdit;
    CPTCREDIT  : THValComboBox;
    CPTDEBIT   : THValComboBox;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure InitMsgBox;

    procedure E_AUXILIAIREDblClick(Sender: TObject);
    procedure E_DEVISEChange(Sender: TObject);
    procedure JOURNALChange(Sender: TObject);
    procedure H_JOURNALDblClick(Sender: TObject);
    procedure DATEGENERATIONExit(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BGenereClick(Sender: TObject);
  private
    DEV : RDEVISE;
    J   : REGJAL;
    DateGene,NowFutur : TDateTime;
    GeneCharge : boolean;
    TPIECES   : TList;
    TitreF    : String;
    Qui : tProfilTraitement;
    Auto, bDejaVu : Boolean;
    procedure AttribCollectif;
    procedure RechercheEcritures;
    procedure ChargeAVide;
    procedure GenereRegul;
    function  WhereCrit : String;
    function  WhereDetail ( RL : RLETTR ) : String;
    procedure RemplirLeM ( Var M : RMVT; RL : RLETTR );
    procedure ValideLaRegul;
    procedure AlimXX_WHERESEL;
    procedure InitCriteres;
    procedure CouvreRegul ( O : TOBM; RL : RLETTR );
    procedure TripoteTitres;
  public
    Regul : PG_LETTRAGE;
  end ;

Implementation
uses
  {$IFDEF MODENT1}
  CPProcGen,
  CPVersion,
  {$ENDIF MODENT1}
  {$IFDEF eAGLCLIENT}
  MenuOLX
  {$ELSE}
  MenuOLG
  {$ENDIF eAGLCLIENT}
    ;

procedure RegulLettrage(Regul,Convert : Boolean );
var
  szArg : String;
begin
  if PasCreerDate(V_PGI.DateEntree) then Exit;
  if _Blocage(['nrCloture','nrBatch','nrSaisieModif','nrLettrage','nrEnca','nrDeca'],True,'nrBatch') then Exit;

  if Regul   then szArg := 'X;'         else szArg := '-;';
  if Convert then szArg := szArg + 'X;' else szArg := szArg + '-;';
  szArg := szArg + '-;';  // Auto = False
  szArg := szArg + 'PRAUCUN;';  // Qui
  AGLLanceFiche('CP','CPLETREGUL','','', szArg);
  _Bloqueur('nrBatch',False);
end ;

procedure RegulInverse;
begin
  if PasCreerDate(V_PGI.DateEntree) then Exit;
  if _Blocage(['nrCloture','nrBatch','nrSaisieModif','nrLettrage','nrEnca','nrDeca'],True,'nrBatch') then Exit;
  AGLLanceFiche('CP','CPLETREGUL','','', '');
  _Bloqueur('nrBatch',False);
end ;

procedure RegulLettrageMP(Regul,Convert : Boolean ; Qui : tProfilTraitement);
var
  szArg : String;
begin
  if PasCreerDate(V_PGI.DateEntree) then Exit;
  if Regul   then szArg := 'X;'         else szArg := '-;';
  if Convert then szArg := szArg + 'X;' else szArg := szArg + '-;';
  szArg := szArg + '-;';  // Auto = False
  szArg := szArg + PTToStr(Qui);  // Qui
  AGLLanceFiche('CP','CPLETREGUL','','', szArg);
end ;

procedure RegulLettrageMPAuto(Regul,Convert : Boolean);
var
  szArg : String;
begin
  if PasCreerDate(V_PGI.DateEntree) then Exit;
  if Regul   then szArg := 'X;'         else szArg := '-;';
  if Convert then szArg := szArg + 'X;' else szArg := szArg + '-;';
  szArg := szArg + 'X;';  // Auto = True
  szArg := szArg + 'PRAUCUN;';  // Qui
  AGLLanceFiche('CP','CPLETREGUL','','', szArg);
end ;

function StrToPT(psz : String) : tProfilTraitement;
begin
  if UpperCase(psz) = 'PRCLIENT'        then Result := prClient else
  if UpperCase(psz) = 'PRFOURNISSEUR'   then Result := prFournisseur else
  if UpperCase(psz) = 'PRETABLISSEMENT' then Result := prEtablissement
                                        else Result := prAucun;
end;

function PTToStr(pPT : tProfilTraitement) : String;
begin
  if pPT = prClient        then Result := 'PRCLIENT' else
  if pPt = prFournisseur   then Result := 'PRFOURNISSEUR' else
  if pPT = prEtablissement then Result := 'PRETABLISSEMENT'
                           else Result := 'PRAUCUN';
end;

procedure TOF_CPLETREGUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPLETREGUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPLETREGUL.OnUpdate ;
begin
  Inherited ;
  {JP 31/05/07 : Code déplacé dans le OnLoad, car en CWas avec les derniers agl, cela posait des problèmes
  if ((VH^.ExoV8.Code<>'') and (StrToDate(GetControlText('E_DATECOMPTABLE'))<VH^.ExoV8.Deb)) then
    SetControlText('E_DATECOMPTABLE', DateToStr(VH^.ExoV8.Deb));
  RechercheEcritures;

  JP 10/06/07 : Par contre, la gestion de la grille, il faut la faire après le chargement,
                en particulier en eAGL}
  {$IFDEF EAGLCLIENT}
  GereSelectionsGrid(GR,QR.TQ);
  {$ELSE}
  GereSelectionsGrid(GR,QR);
  {$ENDIF}
end ;

procedure TOF_CPLETREGUL.OnLoad ;
begin
  Inherited ;
  {JP 31/05/07 : Code déplacé depuis le OnUpdate, car en CWas avec les derniers agl, cela posait des problèmes}
  if ((VH^.ExoV8.Code<>'') and (StrToDate(GetControlText('E_DATECOMPTABLE'))<VH^.ExoV8.Deb)) then
    SetControlText('E_DATECOMPTABLE', DateToStr(VH^.ExoV8.Deb));
  RechercheEcritures;
end ;

procedure TOF_CPLETREGUL.OnArgument (S : String ) ;
var
  E_DEVISE : THValComboBox;
  JOURNAL  : THValComboBox;
  bRegul, bConvert : Boolean;
begin
  Inherited;
  bDejaVu := False;

  // Création des contrôles
  HM := THMsgBox.Create(FMenuG);
  InitMsgBox;
  QR := TFMul(Ecran).Q;
  ModeOppose := TCheckbox(GetControl('MODEOPPOSE', True));
  MAXDEBIT   := THNumEdit(GetControl('MAXDEBIT', True));
  MAXCREDIT  := THNumEdit(GetControl('MAXCREDIT', True));
  CPTCREDIT  := THValComboBox(GetControl('CPTCREDIT', True));
  CPTDEBIT   := THValComboBox(GetControl('CPTDEBIT', True));
  E_DEVISE   := THValComboBox(GetControl('E_DEVISE', True));
  JOURNAL    := THValComboBox(GetControl('JOURNAL', True));

  // Interdit le paramètrage de la liste
  SetControlVisible('BPARAMLISTE', False);

  // Récupère les arguments
  if (S = '')               then begin
    Regul:=pglInverse; TitreF:='REGULINVERSE';
    Qui := prAucun;
    Auto := False;
    end
  else begin
    bRegul := (ReadTokenSt(S) = 'X');
    bConvert := (ReadTokenSt(S) = 'X');
    if bRegul   then begin Regul:=pglRegul;   TitreF:='REGULLETTRAGE'; end else
    if bConvert then begin Regul:=pglConvert; TitreF:='REGULCONVERT'; end else
                     begin Regul:=pglEcart;   TitreF:='ECARTCHANGE'; end;
    if (ReadTokenSt(S) = 'X') then Auto := True else Auto := False;
    Qui := StrToPT(ReadTokenSt(S));
  end;

  // Evénements des contrôles
  THEdit(GetControl('E_AUXILIAIRE', True)).OnDblClick := E_AUXILIAIREDblClick;
  THEdit(GetControl('E_AUXILIAIRE_', True)).OnDblClick := E_AUXILIAIREDblClick;
  THValComboBox(GetControl('E_DEVISE', True)).OnChange := E_DEVISEChange;
  THValComboBox(GetControl('JOURNAL', True)).OnChange := JOURNALChange;
  THLabel(GetControl('H_JOURNALC', True)).OnDblClick := H_JOURNALDblClick;
  THLabel(GetControl('H_JOURNALR', True)).OnDblClick := H_JOURNALDblClick;
  THLabel(GetControl('H_JOURNALE', True)).OnDblClick := H_JOURNALDblClick;
  THEdit(GetControl('DATEGENERATION', True)).OnEnter := DATEGENERATIONExit;
  TToolBarButton97(GetControl('BOUVRIR',True)).OnClick := BValiderClick;
  TToolBarButton97(GetControl('BGENERE',True)).OnClick := BGenereClick;
{$IFDEF EAGLCLIENT}
  GR := THGrid(TFMul(Ecran).FListe);
  SetControlVisible('BSELECTALL', False);
{$ELSE}
  GR := THDBGrid(TFMul(Ecran).FListe);
{$ENDIF}

  // FormCreate
  TPIECES:=TList.Create;

  // FormShow
  InitTablesLibresTiers(TTabSheet(GetControl('PAVANCE')));
  GeneCharge:=True;
  if VH^.TenueEuro then ModeOppose.Caption:=RechDom('TTDEVISETOUTES',V_PGI.DeviseFongible,False);
  if Regul in [pglConvert,pglInverse] then begin
    SetControlVisible('MODEOPPOSE', False);
    SetControlVisible('EXPLIEURO', False);
  end;
  if Regul in [pglConvert,pglInverse] then SetControltext('EXPLIEURO', HM.Mess[20]);
  FillChar(DEV,Sizeof(DEV),#0);
  FillChar(J,Sizeof(J),#0);
  if Regul=pglRegul then begin
    Ecran.Caption:=HM.Mess[0];
    JOURNAL.DataType := 'ttJalRegul';
    SetControlVisible('H_JOURNALR', True);
    Ecran.HelpContext:=7520000;
    SetControlText('REFREGUL', HM.Mess[28]);
    SetControlText('LIBREGUL', HM.Mess[28]);
    E_DEVISE.DataType:='ttDevise';
    E_DEVISE.Value:=V_PGI.DevisePivot;
    QR.Liste:='REGULLETEURO';
    end
  else if Regul=pglEcart then begin
    Ecran.Caption:=HM.Mess[1];
    JOURNAL.DataType := 'ttJalEcart';
    SetControlVisible('H_JOURNALE', True);
    Ecran.HelpContext:=7523000;
    MaxDebit.Visible:=False ;
    SetControlVisible('H_SD', False);
    MaxCredit.Visible:=False;
    SetControlVisible('H_SC', False);
    SetControlText('REFREGUL', HM.Mess[29]);
    SetControlText('LIBREGUL', HM.Mess[29]);
    E_DEVISE.DataType := 'ttDeviseEtat';
    E_DEVISE.Value := '';
    if E_DEVISE.Items.Count>0 then begin
      E_DEVISE.ItemIndex:=0;
      E_DEVISEChange(Nil);
    end;
    QR.Liste:='REGULLETEURO';
    end
  else begin{Convert,Inverse}
    Ecran.Caption:=HM.Mess[18];
    JOURNAL.DataType := 'ttJournal';
    JOURNAL.Enabled := False;
    SetControlVisible('H_JOURNALC', True);
//    JOURNAL.Value := VH^.JalEcartEuro;
    MaxDebit.Visible:=False ;
    SetControlVisible('H_SD', False);
    MaxCredit.Visible:=False;
    SetControlVisible('H_SC', False);
    SetControlProperty('BOUVRIR', 'HINT', HM.Mess[21]);
    Ecran.HelpContext:=7520500;
    SetControlText('REFREGUL', HM.Mess[30]);
    SetControlText('LIBREGUL', HM.Mess[30]);
    E_DEVISE.DataType:='ttDevise';
    E_DEVISE.Value:=V_PGI.DevisePivot;
    QR.Liste:='ECARTCONVERTEURO';
    If Regul=pglInverse Then begin
      If VH^.TenueEuro Then QR.Titres:='Général;Auxiliaire;Intitulé auxiliaire;Nb D;Débit Euro;Débit Frcs;Nb C; ;Crédit Frcs;Solde Euro;Solde Frcs;Let;Dev;LD;LE;';
    end;
  end;

  SetControlText('DATEGENERATION', DateToStr(V_PGI.DateEntree));
  DateGene:=V_PGI.DateEntree;
  InitCriteres;
  if TFMul(Ecran).FFiltres.Text='DEFAUT' then
    TFMul(Ecran).BChercheClick(Nil)
  else begin
    ChargeAVide;
    SetFocusControl('E_AUXILIAIRE');
    if ((Regul<>pglConvert) and (Regul<>pglInverse)) then begin
      if JOURNAL.Items.Count>0 then begin
        JOURNAL.ItemIndex:=0;
        JournalChange(Nil);
      end;
      end
    else begin
//      Journal.Value:=VH^.JalEcartEuro;
    end;
  end;
  GeneCharge:=False;
  UpdateCaption(Ecran);
  Tripotetitres;
  If Auto Then begin
    TFMul(Ecran).BChercheClick(Nil);
    If Not QR.Eof Then TFMul(Ecran).bSelectAllClick(Nil);
  end;

end ;

procedure TOF_CPLETREGUL.OnClose ;
var
  Ind : integer;
begin
  Inherited ;
  if TPieces.Count>0 then begin
    if Regul=pglRegul then Ind:=3 else if Regul=pglEcart then Ind:=4 else Ind:=14;
    if not bDejaVu then
      if HM.Execute(13,Ecran.Caption,'')=mrYes then VisuPiecesGenere(TPieces,EcrGen,Ind);
    VideListe(TPieces);
    TPieces.Free;
  end;
  if (Qui <> prAucun) then
    if (Not MonProfilOk(Qui)) And (Not Auto) Then _Bloqueur('nrBatch',False);
end ;

procedure TOF_CPLETREGUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPLETREGUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPLETREGUL.InitMsgBox;
begin
HM.Mess.Add('Régularisations de lettrage');
HM.Mess.Add('Différences de change');
HM.Mess.Add('2;?caption?;Vous n''avez rien sélectionné.;E;O;O;O;');
HM.Mess.Add('3;?caption?;Vous devez choisir un journal de régularisation.;W;O;O;O;');
HM.Mess.Add('4;?caption?;Vous devez renseigner les comptes de régularisation.;W;O;O;O;');
HM.Mess.Add('5;?caption?;La date de génération que vous avez saisie est incorrecte.;W;O;O;O;');
HM.Mess.Add('6;?caption?;La date de génération est sur un exercice non ouvert.;W;O;O;O;');
HM.Mess.Add('7;?caption?;La date de génération est sur un exercice non ouvert.;W;O;O;O;');
HM.Mess.Add('8;?caption?;La date de génération est antérieure à la clôture provisoire.;W;O;O;O;');
HM.Mess.Add('9;?caption?;La date de génération est antérieure à la clôture définitive.;W;O;O;O;');
HM.Mess.Add('10;?caption?;Le compte de régularisation débit n''existe pas.;W;O;O;O;');
HM.Mess.Add('11;?caption?;Le compte de régularisation crédit n''existe pas.;W;O;O;O;');
HM.Mess.Add('12;?caption?;Confirmez-vous la génération des pièces de régularisation ?;Q;YN;Y;Y;');
HM.Mess.Add('13;?caption?;Voulez-vous voir la liste des pièces générées ?;Q;YN;N;N;');
HM.Mess.Add('14;?caption?;Confirmez-vous la génération des pièces d''écart de change ?;Q;YN;Y;Y;');
HM.Mess.Add('Régularisation non effectuée. Certaines écritures sont en traitement par un autre utilisateur.');
HM.Mess.Add('16;?caption?;La régularisation est impossible, le journal ne possède pas de facturier.;W;O;O;O;');
HM.Mess.Add('17;');
HM.Mess.Add('Ecarts de conversion');
HM.Mess.Add('19;?caption?;Confirmez-vous la génération des pièces d''écart de conversion ?;Q;YN;Y;Y;');
HM.Mess.Add('Devise de l''écart');
HM.Mess.Add('Générer les pièces d''écart de conversion');
HM.Mess.Add('Débit Euro');
HM.Mess.Add('Crédit Euro');
HM.Mess.Add('Solde Euro');
HM.Mess.Add('Débit');
HM.Mess.Add('Crédit');
HM.Mess.Add('Solde');
HM.Mess.Add('Régularisation');
HM.Mess.Add('Différence de change');
HM.Mess.Add('Ecart de conversion');
HM.Mess.Add('31;');
end;

procedure TOF_CPLETREGUL.AlimXX_WHERESEL;
var
  QS : TQuery;
  Gene,Aux,Lett,St,St1 : String;
begin
  St:='';
  St1:='';
  QS := OpenSQL('SELECT * FROM CPMPTEMPOR WHERE CTT_USER="'+V_PGI.User+'" AND CTT_TYPETRAITEMENT="REG" ', True);
  While Not QS.Eof Do begin
    Gene := QS.FindField('CTT_GENERAL').AsString;
    Aux  := QS.FindField('CTT_AUXILIAIRE').AsString;
    Lett := QS.FindField('CTT_ZONELIBRE1').AsString;
    St1:='(E_GENERAL="'+Gene+'" AND E_AUXILIAIRE="'+Aux+'" AND E_LETTRAGE="'+Lett+'") ';
    If St='' Then St:=St1
             Else St:=St+' OR '+St1;
    QS.Next;
  end;
  Ferme(QS);
  If St='' Then Exit;
  St:='E_JOURNAL<>"'+W_W+'" AND ('+St+')';
  SetControlText('XX_WHERESEL', St);
end;

procedure TOF_CPLETREGUL.AttribCollectif;
Var
  Q : TQuery;
begin
  if ((GetControlText('E_AUXILIAIRE_')<>'') or (GetControlText('E_GENERAL')<>'') or (GetControlText('E_GENERAL_')<>'') or (GetControlText('E_AUXILIAIRE')='')) then Exit;
  Q := OpenSQL('SELECT T_COLLECTIF FROM TIERS WHERE T_AUXILIAIRE="'+GetControlText('E_AUXILIAIRE')+'"',True);
  if Not Q.EOF then SetControlText('E_GENERAL', Q.Fields[0].AsString);
  Ferme(Q);
end;

procedure TOF_CPLETREGUL.ChargeAVide;
var
  ind, iFrom : integer;
  St : String;
  szSelect, szFrom : String;
begin
  If Qui<>prAucun Then St:=' LEFT OUTER JOIN GENERAUX ON E_GENERAL=G_GENERAL ';
  if Regul=pglRegul then Ind:=0 else
  if Regul=pglEcart then Ind:=1
                    else Ind:=3;
  szSelect := SelectQDL(Ind);
  System.Delete(szSelect,1,7);
  iFrom := Pos(' FROM ', UpperCase(szSelect));
  szFrom := Copy(szSelect, iFrom, Length(szSelect)-iFrom+1);
  System.Delete(szSelect,iFrom,Length(szFrom));
  System.Delete(szFrom,1,6);
  SetControlText('XX_SELECT', szSelect);
  SetControlText('XX_FROM', szFrom);
  SetControlText('XX_WHEREBASE', LWhereVide);

  If MonProfilOk(Qui) Then begin
{$IFDEF EAGLCLIENT}
    St:=WhereProfilUser(QR.TQ,Qui);
{$ELSE}
    St:=WhereProfilUser(QR,Qui);
{$ENDIF}
    SetControlText('XX_WHERECRIT', St);
  end;
  SetControlText('XX_GROUPBY', 'E_GENERAL, E_AUXILIAIRE, E_DEVISE, E_LETTRAGE');
  ChangeAspectLettre(GR);
end;

procedure TOF_CPLETREGUL.CouvreRegul(O: TOBM; RL: RLETTR);
Var
  TotD,TotF,Taux : double;
begin
  TotF:=O.GetMvt('E_DEBIT')+O.GetMvt('E_CREDIT');
  TotD:=O.GetMvt('E_DEBITDEV')+O.GetMvt('E_CREDITDEV');
  Case Regul of
    pglRegul : begin
                 if RL.LettrageDevise then begin
                   O.PutMvt('E_COUVERTUREDEV',TotD);
                   Taux:=O.GetMvt('E_TAUXDEV');
                   if Arrondi(O.GetMvt('E_COUVERTURE')-TotF,V_PGI.OkDecV)<>0     then O.PutMvt('E_COUVERTURE',DeviseToPivot(TotD,Taux,DEV.Quotite));
                   end
                 else begin
                   O.PutMvt('E_COUVERTURE',TotF);
                   ConvertCouverture(O,tsmPivot);
                 end;
               end;
    pglEcart : begin
                {Ecart de change couvre tout}
                O.PutMvt('E_COUVERTUREDEV',TotD);
                O.PutMvt('E_COUVERTURE',TotF);
               end;
    pglConvert : begin
                   {Ecart de conversion couvre tout}
                   O.PutMvt('E_COUVERTUREDEV',TotD);
                   O.PutMvt('E_COUVERTURE',TotF);
                 end;
    pglInverse : begin
                   {Ecart de conversion couvre tout}
                   O.PutMvt('E_COUVERTUREDEV',TotD);
                   O.PutMvt('E_COUVERTURE',TotF);
                 end;
  end;
end;

procedure TOF_CPLETREGUL.GenereRegul;
var
  RL : RLETTR;
  Q  : TQuery;
  X,LimiteP,LimiteE,Montant,MenOppose : Double;
  St,Cpte,StInd : String;
  O,OOO     : TOBM;
  Premier   : boolean;
  M         : RMVT;
  XX        : T_ECARTCHANG;
  DMin,DMax : TDateTime;
  ForceSaisieEuro : String;
begin
  ForceSaisieEuro:='';
  if ((Regul<>pglConvert) and (Regul<>pglInverse)) then begin
{$IFDEF EAGLCLIENT}
    Montant:=Arrondi(QR.TQ.Fields[7].AsFloat,V_PGI.OkDecV); // Crédit Pivot
{$ELSE}
    Montant:=Arrondi(GR.Fields[7].AsFloat,V_PGI.OkDecV);    // Crédit Pivot
{$ENDIF}
    MenOppose:=0;
    end
  else begin
{$IFDEF EAGLCLIENT}
    Montant:=Arrondi(QR.TQ.Fields[9].AsFloat,V_PGI.OkDecV);    // Crédit FRF
    MenOppose:=Arrondi(QR.TQ.Fields[10].AsFloat,V_PGI.OkDecE); // Solde Euro
{$ELSE}
    Montant:=Arrondi(GR.Fields[9].AsFloat,V_PGI.OkDecV);    // Crédit FRF
    MenOppose:=Arrondi(GR.Fields[10].AsFloat,V_PGI.OkDecE); // Solde Euro
{$ENDIF}
  end;

  DMax:=0;
  DMin:=0;
  LimiteP:=Resolution(V_PGI.OkDecV);
  LimiteE:=Resolution(V_PGI.OkDecE);
  if ((Regul<>pglConvert) and (Regul<>pglInverse)) then begin
    if Abs(Montant)<LimiteP then Exit;
    end
  else begin
    if ((Abs(Montant)<LimiteP) and (Abs(MenOppose)<LimiteE)) then Exit;
  end;
  Premier:=True;
  FillChar(RL,Sizeof(RL),#0);
{$IFDEF EAGLCLIENT}
  RL.General:=QR.TQ.Fields[0].AsString;     // Général
  RL.Auxiliaire:=QR.TQ.Fields[1].AsString;  // Auxiliaire
{$ELSE}
  RL.General:=GR.Fields[0].AsString;     // Général
  RL.Auxiliaire:=GR.Fields[1].AsString;  // Auxiliaire
{$ENDIF}
  RL.Appel:=tlMenu;
  RL.GL:=Nil;

  if ((Regul<>pglConvert) and (Regul<>pglInverse)) then begin
{$IFDEF EAGLCLIENT}
    RL.CodeLettre:=QR.TQ.Fields[8].AsString;     // Crédit Euro
    RL.DeviseMvt:=QR.TQ.Fields[9].AsString ;     // Crédit FRF
    RL.LettrageDevise:=(RL.DeviseMvt<>V_PGI.DevisePivot);
//    RL.LettrageOppose:=(QR.TQ.Fields[11].AsString='X');  // Solde FRF
{$ELSE}
    RL.CodeLettre:=GR.Fields[8].AsString;     // Crédit Euro
    RL.DeviseMvt:=GR.Fields[9].AsString ;     // Crédit FRF
    RL.LettrageDevise:=(RL.DeviseMvt<>V_PGI.DevisePivot);
//    RL.LettrageOppose:=(GR.Fields[11].AsString='X');  // Solde FRF
{$ENDIF}
    end
  else begin
{$IFDEF EAGLCLIENT}
    RL.CodeLettre:=QR.TQ.Fields[11].AsString;     // Solde FRF
    RL.DeviseMvt:=QR.TQ.Fields[12].AsString ;     // DEV
    RL.LettrageDevise:=(RL.DeviseMvt<>V_PGI.DevisePivot);
//    RL.LettrageOppose:=(QR.TQ.Fields[14].AsString='X');  // LE
{$ELSE}
    RL.CodeLettre:=GR.Fields[11].AsString;     // Solde FRF
    RL.DeviseMvt:=GR.Fields[12].AsString ;     // DEV
    RL.LettrageDevise:=(RL.DeviseMvt<>V_PGI.DevisePivot);
//    RL.LettrageOppose:=(GR.Fields[14].AsString='X');  // LE
{$ENDIF}
  end;
  if ((Regul=pglRegul) and (RL.LettrageDevise)) then begin
    {Récupérer l'écart en devise}
    Q := OpenSQL('SELECT SUM(E_DEBITDEV-E_CREDITDEV) FROM Ecriture WHERE '+WhereDetail(RL),True);
    if Not Q.EOF then Montant:=Q.Fields[0].AsFloat;
    Ferme(Q);
  end;

  if Regul=pglEcart then begin
    {Récupérer l'écart opposé aussi}
    Q:=OpenSQL('Select SUM(E_DEBIT-E_CREDIT) from Ecriture Where '+WhereDetail(RL),True);
    if Not Q.EOF then MenOppose:=Arrondi(Q.Fields[0].AsFloat,V_PGI.OkDecE);
    Ferme(Q);
  end;

  if ((Regul<>pglConvert) and (Regul<>pglInverse)) then begin
    if Montant=0 then Exit;
    end
  else begin
    if ((Montant=0) and (MenOppose=0)) then Exit;
  end;

  if Regul in [pglConvert,pglInverse] then begin
    X := MenOppose;
    MenOppose := Montant;
    Montant := X;
  end;

  Q := OpenSQL('Select * from Ecriture Where '+WhereDetail(RL),True);
  While Not Q.EOF do begin
    O := TOBM.Create(EcrGen,'',False);
    O.ChargeMvt(Q);
    O.PutMvt('E_ETATLETTRAGE','TL');
    O.PutMvt('E_LETTRAGE',uppercase(O.GetMvt('E_LETTRAGE')));
    {Etat}
    St := O.GetMvt('E_ETAT');
    if ((Montant>0) or ((Montant=0) and (MenOppose>0))) then StInd:=IntToStr(CPTDEBIT.ItemIndex+1)
                                                        else StInd:=IntToStr(CPTCREDIT.ItemIndex+1);
    if Length(St)>=4 then begin
      if Regul=pglRegul then St[1]:='X'
      else begin
        if Regul=pglEcart then St[2]:='X' else
          if Regul=pglConvert then St[2]:='#' else St[2]:='&';
        St[3]:='-';
      end;
      St[4]:=StInd[1];
    end;
    O.PutMvt('E_ETAT',St);
    CouvreRegul(O,RL);
    {Dates}
    if Premier then begin
      DMin:=O.GetMvt('E_DATEPAQUETMIN'); if DateGene<DMin then DMin:=DateGene;
      DMax:=O.GetMvt('E_DATEPAQUETMAX'); if DateGene>DMax then DMax:=DateGene;
    end;
    O.PutMvt('E_DATEPAQUETMIN',DMin);
    O.PutMvt('E_DATEPAQUETMAX',DMax);
    if Not GoReqMajLet(O,RL.General,RL.Auxiliaire,NowFutur,True) then V_PGI.IoError:=oeUnknown;
    if Premier then begin
      RemplirLeM(M,RL);
      if ((Regul<>pglConvert) and (Regul<>pglInverse)) then begin
        if Montant>0 then Cpte:=CPTDEBIT.Text
                     else Cpte:=CptCredit.Text;
        end
      else begin{Convert,Inverse}
        if (Montant>0) or ((Montant=0) and (MenOppose>0)) then Cpte:=CPTDEBIT.Text
                                                          else Cpte:=CptCredit.Text;
      end;
      XX.Cpte:=Cpte;
      XX.Regul:=Regul;
      XX.Ref:=GetControlText('REFREGUL');
      XX.Lib:=GetControlText('LIBREGUL');
      XX.Quotite:=DEV.Quotite;
      XX.Cpte:=Cpte;
      XX.DPMin:=DMin;
      XX.DPMax:=DMax;
      XX.Decimale:=DEV.Decimale;
      Case Regul of
         pglRegul   : begin XX.Montant1:=Montant; XX.Montant2:=0; end;
         pglEcart   : begin XX.Montant1:=Montant; XX.Montant2:=MenOppose; end;
         pglConvert : begin XX.Montant1:=Montant; XX.Montant2:=MenOppose; end;
         pglInverse : begin XX.Montant1:=Montant; XX.Montant2:=MenOppose; ForceSaisieEuro:='-'; end;
      end;
      OOO:=CreerPartieDoubleLett(M,RL,XX,O,ForceSaisieEuro);
      TPieces.Add(OOO);
    end;
    Q.Next;
    Premier:=False;
    O.Free;
  end;
  Ferme(Q);
end;

procedure TOF_CPLETREGUL.InitCriteres;
var
  DD : TDateTime;
begin
  If Auto Then begin
{$IFNDEF EAGLCLIENT}
    SetControlVisible('BSELECTALL', True);
{$ENDIF}
    If GetControlText('XX_WHERESEL') ='' Then begin
      AlimXX_WHERESEL;
      SetControlText('E_GENERAL', '');
      SetControlText('E_GENERAL_', '');
      SetControlText('E_AUXILIAIRE', '');
      SetControlText('E_AUXILIAIRE_', '');
      SetControlText('DATEPAQUETMIN', StDate1900);
      SetControlText('DATEPAQUETMAX', StDate2099);
      SetControlText('E_DEVISE', V_PGI.DevisePivot);
      THValComboBox(GetControl('T_NATUREAUXI', True)).ItemIndex := 0;
      ModeOppose.Checked:=FALSE;
      SetControlText('T_TABLE0', '');
      SetControlText('T_TABLE1', '');
      SetControlText('T_TABLE2', '');
      SetControlText('T_TABLE3', '');
      SetControlText('T_TABLE4', '');
      SetControlText('T_TABLE5', '');
      SetControlText('T_TABLE6', '');
      SetControlText('T_TABLE7', '');
      SetControlText('T_TABLE8', '');
      SetControlText('T_TABLE9', '');
    end;
    end
  eLse begin
    if VH^.Precedent.Code<>'' then SetControlText('DATEPAQUETMIN', DateToStr(VH^.Precedent.Deb))
                              else SetControlText('DATEPAQUETMIN', DateToStr(VH^.Encours.Deb));
    SetControlText('DATEPAQUETMAX', DateToStr(V_PGI.DateEntree));
    if ((VH^.ExoV8.Code<>'') and (StrToDate(GetControlText('E_DATECOMPTABLE'))<VH^.ExoV8.Deb)) then SetControlText('E_DATECOMPTABLE', DateToStr(VH^.ExoV8.Deb));
  end;
  DD := StrToDate(GetControlText('DATEGENERATION'));
  if ModeRevisionActive(DD) then begin
    SetControlText('DATEGENERATION', DateToStr(V_PGI.DateEntree));
    DateGene:=V_PGI.DateEntree;
  end;
end;

procedure TOF_CPLETREGUL.RechercheEcritures;
var
  StXP,StXN,StXP2,StXN2,StXP3,StXN3 : String;
  iFrom, Ind : integer;
  szSelect, szFrom, szHaving : String;
begin
  GR.ClearSelected;

  if Regul=pglRegul then Ind:=1 else
  if Regul=pglEcart then Ind:=0 else
                         Ind:=3;
  szSelect := SelectQDL(Ind);
  System.Delete(szSelect,1,7);
  iFrom := Pos(' FROM ', UpperCase(szSelect));
  szFrom := Copy(szSelect, iFrom, Length(szSelect)-iFrom+1);
  System.Delete(szSelect,iFrom,Length(szFrom));
  System.Delete(szFrom,1,6);
  SetControlText('XX_SELECT', szSelect);
  SetControlText('XX_FROM', szFrom);
  SetControlText('XX_WHEREBASE', LWhereBase(Regul=pglRegul,True,False,False));

  if Regul=pglRegul then SetControlText('E_ETATLETTRAGE', 'PL')
                    else SetControlText('E_ETATLETTRAGE', 'TL');
  if (EstSpecif('51188')) And (Regul=pglConvert) And (Not (ctxPCL in V_PGI.PGIContexte)) then begin
    SetControlText('E_DATEPAQUETMIN', UsDateTime(EncodeDate(2002,01,01)));
  end;
  SetControlText('XX_WHERECRIT', WhereCrit);
  if Regul=pglRegul then begin
    if GetControlText('E_DEVISE') = V_PGI.DevisePivot then begin
      szHaving := '(SUM(E_DEBIT-E_CREDIT)>0 AND SUM(E_DEBIT-E_CREDIT)<'+StrFPoint(MAXDEBIT.Value)+')'
                 +'OR (SUM(E_DEBIT-E_CREDIT)<0 AND SUM(E_DEBIT-E_CREDIT)>'+StrFPoint(-MAXCREDIT.Value)+')';
      end
    else begin
      szHaving := '(SUM(E_DEBITDEV-E_CREDITDEV)>0 AND SUM(E_DEBITDEV-E_CREDITDEV)<'+StrFPoint(MAXDEBIT.Value)+')'
                 +'OR (SUM(E_DEBITDEV-E_CREDITDEV)<0 AND SUM(E_DEBITDEV-E_CREDITDEV)>'+StrFPoint(-MAXCREDIT.Value)+')';
    end;
    end
  else if Regul in [pglConvert,pglInverse] then begin
    StXP:=StrFPoint(9*Resolution(V_PGI.OkDecV+1)) ;
    StXN:=StrFPoint(-9*Resolution(V_PGI.OkDecV+1));
    StXP2:=StrFPoint(9*Resolution(V_PGI.OkDecE+1));
    StXN2:=StrFPoint(-9*Resolution(V_PGI.OkDecE+1));
    StXP3:=StrFPoint(9*Resolution(DEV.Decimale+1));
    StXN3:=StrFPoint(-9*Resolution(DEV.Decimale+1));
    (*
    if EstMonnaieIN(DEV.Code) then  begin
      szHaving := '(Sum(E_COUVERTURE*'+SoldeSurDC+') Not Between '+StXN+' AND '+StXP+')) AND ('
                 +'Sum(E_COUVERTUREDEV*'+SoldeSurDC+') Between '+StXN3+' AND '+StXP3+')';
      end
    else *)
    begin
      if Regul=pglConvert then begin
        if VH^.TenueEuro then begin
          szHaving := '(Sum(E_COUVERTURE*'+SoldeSurDC+') Between '+StXN+' AND '+StXP+')';
          end
        else begin
          szHaving := '(Sum(E_COUVERTURE*'+SoldeSurDC+') Not Between '+StXN2+' AND '+StXP2+')';
        end;
        end
      else if Regul=pglInverse then begin
        if VH^.TenueEuro then begin
          szHaving := '(Sum(E_COUVERTURE*'+SoldeSurDC+') Not Between '+StXN+' AND '+StXP+')';
          end
        else begin
          szHaving := '(Sum(E_COUVERTURE*'+SoldeSurDC+') Between '+StXN2+' AND '+StXP2+')';
        end;
      end;
    end;
    end
  else begin{écart change}
    if VH^.TenueEuro then begin
      StXP:=StrFPoint(9*Resolution(V_PGI.OkDecV+1));
      StXN:=StrFPoint(-9*Resolution(V_PGI.OkDecV+1));
      szHaving := 'Sum(E_COUVERTURE*'+SoldeSurDC+') Not Between '+StXN+' AND '+StXP;
      end
    else begin
      StXP:=StrFPoint(9*Resolution(V_PGI.OkDecE+1)); StXN:=StrFPoint(-9*Resolution(V_PGI.OkDecE+1));
//      szHaving := 'Sum(E_COUVERTUREEURO*'+SoldeSurDC+') Not Between '+StXN+' AND '+StXP;
    end;
  end;
  SetControlText('XX_HAVING', szHaving);
  SetControlText('XX_GROUPBY', 'E_GENERAL, E_AUXILIAIRE, E_DEVISE, E_LETTRAGE');
  ChangeAspectLettre(GR);

(* JP 12/06/07 : Déplacé dans le OnUpdate
{$IFDEF EAGLCLIENT}
  GereSelectionsGrid(GR,QR.TQ);
{$ELSE}
  GereSelectionsGrid(GR,QR);
{$ENDIF}
*)
  TripoteTitres;
end;

procedure TOF_CPLETREGUL.RemplirLeM(var M: RMVT; RL: RLETTR);
var
  Facturier : String3;
begin
  FillChar(M,Sizeof(M),#0);
  M.Jal := GetControlText('JOURNAL');
  Facturier:=J.Facturier;
  M.ModeSaisieJal:=J.ModeSaisie;
//SetIncNum(EcrGen,Facturier,M.Num);
  M.Axe:='';
  M.Etabl:=VH^.EtablisDefaut;
  if ((Regul=pglRegul) and (RL.LettrageDevise)) or (Regul in [pglConvert,pglInverse]) then begin
    M.CodeD:=DEV.Code;
    M.TauxD:=DEV.Taux;
    end
  else begin
    M.CodeD:=V_PGI.DevisePivot;
    M.TauxD:=1;
  end;
  M.DateTaux:=V_PGI.DateEntree;
  M.Valide:=False;
  M.Simul:='N';
  if Regul in [pglRegul,pglConvert,pglInverse] then M.Nature:='OD' else M.Nature:='ECC';
  M.DateC:=DateGene;
  M.Exo:=QuelExo(DateToStr(M.DateC));
  M.General:='';
  SetIncNum(EcrGen,Facturier,M.Num,M.DateC);
end;

procedure TOF_CPLETREGUL.TripoteTitres;
var
  DV : RDEVISE;
begin

  if Regul=pglEcart then begin
{$IFDEF EAGLCLIENT}
    GR.Cells[4,0] := HM.Mess[22];
    GR.Cells[6,0] := HM.Mess[23];
    GR.Cells[7,0] := HM.Mess[24];
{$ELSE}
    GR.Columns.Items[4].Title.Caption := HM.Mess[22];
    GR.Columns.Items[6].Title.Caption := HM.Mess[23];
    GR.Columns.Items[7].Title.Caption := HM.Mess[24];
{$ENDIF}
    end
  else if Regul in [pglConvert,pglInverse] then begin
    if VH^.TenueEuro then begin
      DV.Code:=V_PGI.DeviseFongible;
      GetInfosDevise(DV);
      if DV.Symbole='' then DV.Symbole:=Copy(DV.Libelle,1,2);
{$IFDEF EAGLCLIENT}
      GR.Cells[6,0]  := HM.Mess[25]+' '+DV.Symbole;
      GR.Cells[9,0]  := HM.Mess[26]+' '+DV.Symbole;
      GR.Cells[11,0] := HM.Mess[27]+' '+DV.Symbole;
{$ELSE}
//    if (GR.Columns.Count >= 11) then begin
      GR.Columns.Items[6].Title.Caption:=HM.Mess[25]+' '+DV.Symbole;
      GR.Columns.Items[9].Title.Caption:=HM.Mess[26]+' '+DV.Symbole;
      GR.Columns.Items[11].Title.Caption:=HM.Mess[27]+' '+DV.Symbole;
//    end;
{$ENDIF}
    end;
  end;
{$IFDEF EAGLCLIENT}
  TFMul(Ecran).HMTrad.ResizeGridColumns(GR);
{$ELSE}
  TFMul(Ecran).HMTrad.ResizeDBGridColumns(GR);
{$ENDIF}
end;

procedure TOF_CPLETREGUL.ValideLaRegul;
var
  i : integer;
begin
  InitMove(GR.NbSelected,'');
  for i:=0 to GR.NbSelected-1 do begin
    GR.GotoLeBookMark(i);
{$IFDEF EAGLCLIENT}
    QR.TQ.Seek(GR.Row-1);
{$ENDIF}
    MoveCur(FALSE);
    if V_PGI.IoError=oeOK then GenereRegul
                          else Break;
  end;
  FiniMove;
end;

function TOF_CPLETREGUL.WhereCrit: String;
var
  St : String;
begin
  St:=' E_DATEPAQUETMIN>="'+USDATETIME(StrToDate(GetControlText('DATEPAQUETMIN')))+'"'
     +' AND E_DATEPAQUETMAX<="'+USDATETIME(StrToDate(GetControlText('DATEPAQUETMAX')))+'"';
  Case Regul of
     pglRegul : if GetControlText('E_DEVISE') = V_PGI.DevisePivot then St:=St+' AND E_LETTRAGEDEV="-"' else
                if GetControlText('E_DEVISE') <> '' then St:=St+' AND E_LETTRAGEDEV="X"';
     pglEcart : begin
                St:=St+' AND E_LETTRAGEDEV="X"';
                St:=St+' AND (E_DATEPAQUETMIN>="'+USDATETIME(V_PGI.DateDebutEuro)+'"'
                      +' OR E_DATEPAQUETMAX<"'+USDATETIME(V_PGI.DateDebutEuro)+'")';
                end;
     pglConvert :;
  end;
  Result := St;
end;

function TOF_CPLETREGUL.WhereDetail(RL: RLETTR): String;
var
  EL : String;
begin
  if Regul=pglRegul then EL:='PL'
                    else EL:='TL';
  Result:=' E_AUXILIAIRE="'+RL.Auxiliaire+'"'+' AND E_GENERAL="'+RL.General+'"'
         +' AND E_ETATLETTRAGE="'+EL+'" AND E_LETTRAGE="'+RL.CodeLettre+'"';
end;

procedure TOF_CPLETREGUL.E_AUXILIAIREDblClick(Sender: TObject);
begin
  AttribCollectif;
end;

procedure TOF_CPLETREGUL.E_DEVISEChange(Sender: TObject);
var
  DEVF : RDEVISE;
begin
  if GetControlText('E_DEVISE') = DEV.Code then Exit;
  DEV.Code := GetControlText('E_DEVISE');
  GetInfosDevise(DEV);
  DEV.Taux:=GetTaux(DEV.Code,DEV.DateTaux,V_PGI.DateEntree);
  ChangeMask(MAXDEBIT,DEV.Decimale,DEV.Symbole);
  ChangeMask(MAXCREDIT ,DEV.Decimale,DEV.Symbole);
  if ((VH^.TenueEuro) and (DEV.Code=V_PGI.DevisePivot) and (ModeOppose.Checked)) then begin
    DEVF.Code:=V_PGI.DeviseFongible;
    GetInfosDevise(DEVF);
    end
  else begin
    DEVF:=DEV;
  end;

  MAXDEBIT.Value:=DEVF.MaxDebit;
  MAXCREDIT.Value:=DEVF.MaxCredit;
  if Regul=pglEcart then begin
    CptDebit.Items.Clear;
    CptCredit.Items.Clear;
    CptDebit.Items.Add(DEV.CptDebit);
    CptCredit.Items.Add(DEV.CptCredit);
    if CptDebit.Items.Count>0 then CptDebit.ItemIndex:=0;
    if CptCredit.Items.Count>0 then CptCredit.ItemIndex:=0;
    CPTDEBIT.Enabled:=False; CPTCREDIT.Enabled:=False;
  end;
end;

procedure TOF_CPLETREGUL.JOURNALChange(Sender: TObject);
var
  i : integer;
begin
  if GetControlText('JOURNAL') = J.Journal then Exit;
  RemplirInfosRegul(J,GetControlText('JOURNAL'));
  if ((GetControlText('JOURNAL')<>'') and (J.Facturier='') and (Regul<>pglConvert) and (Regul<>pglInverse)) then begin
    if Not GeneCharge then HM.Execute(16,Ecran.Caption,'');
    SetControlText('JOURNAL', '');
    Exit;
  end;
  if Regul=pglRegul then begin
    CptDebit.Items.Clear;
    CptCredit.Items.Clear;
    for i:=1 to 3 do
      if J.D[i]<>'' then CptDebit.Items.Add(J.D[i]);
    for i:=1 to 3 do
      if J.C[i]<>'' then CptCredit.Items.Add(J.C[i]);
    if CptDebit.Items.Count>0 then CptDebit.ItemIndex:=0;
    if CptCredit.Items.Count>0 then CptCredit.ItemIndex:=0;
    end
  else if Regul in [pglConvert,pglInverse] then begin
    CptDebit.Items.Clear;
    CptCredit.Items.Clear;
//    CptDebit.Items.Add(VH^.EccEuroDebit);
//    CptCredit.Items.Add(VH^.EccEuroCredit);
    if CptDebit.Items.Count>0 then CptDebit.ItemIndex:=0;
    if CptCredit.Items.Count>0 then CptCredit.ItemIndex:=0;
  end;
end;

procedure TOF_CPLETREGUL.H_JOURNALDblClick(Sender: TObject);
var
  A : TActionFiche;
begin
  if Not ExJaiLeDroitConcept(TConcept(ccJalModif),False) then A:=taConsult else A:=taModif;
  if GetControlText('JOURNAL') <> '' then begin
    FicheJournal(Nil,'', GetControlText('JOURNAL'),A,0);
    if A=taModif then JournalChange(Nil);
  end;
end;

procedure TOF_CPLETREGUL.DATEGENERATIONExit(Sender: TObject);
var
  DD : TDateTime;
  Err : integer;
begin
  if Not IsValidDate(GetControlText('DATEGENERATION')) then begin
    HM.Execute(5,Ecran.Caption,'');
    SetControlText('DATEGENERATION', DateToStr(V_PGI.DateEntree));
    DateGene:=V_PGI.DateEntree;
    end
  else begin
    DD := StrToDate(GetControlText('DATEGENERATION'));
    Err := DateCorrecte(DD);
    if Err>0 then begin
      HM.Execute(5+Err,Ecran.Caption,'');
      SetControlText('DATEGENERATION', DateToStr(V_PGI.DateEntree));
      DateGene:=V_PGI.DateEntree;
      end
    else begin
      if RevisionActive(DD) then begin
        SetControlText('DATEGENERATION', DateToStr(V_PGI.DateEntree));
        DateGene := V_PGI.DateEntree;
        end
      else begin
        DateGene:=DD;
      end;
    end;
  end;
end;

procedure TOF_CPLETREGUL.BValiderClick(Sender: TObject);
Var
  Num : integer;
begin
  NowFutur:=NowH;
  if GetControlText('JOURNAL') = '' then begin
    HM.Execute(3,Ecran.Caption,'');
    Exit;
  end;
  if ((CPTDEBIT.Text='') or (CPTCREDIT.Text='')) then begin
    HM.Execute(4,Ecran.Caption,'');
    Exit;
  end;
  if Not Presence('GENERAUX','G_GENERAL',CPTDEBIT.Text) then begin
    HM.Execute(10,Ecran.Caption,'');
    Exit;
  end;

  if Not Presence('GENERAUX','G_GENERAL',CPTCREDIT.Text) then begin
    HM.Execute(11,Ecran.Caption,'');
    Exit;
  end;

  if RevisionActive(DateGene) then Exit;

  if GR.NbSelected<=0 then begin
    HM.Execute(2,Ecran.Caption,'');
    Exit;
  end;

  if Regul=pglRegul then Num:=12
                    else if Regul=pglEcart then Num:=14 else Num:=19;
  if HM.Execute(Num,Ecran.Caption,'')<>mrYes then Exit;
  if Transactions(ValideLaRegul,5)<>oeOK then MessageAlerte(HM.Mess[15]);
  bDejaVu := False;
  TFMul(Ecran).BChercheClick(Nil);
end;

procedure TOF_CPLETREGUL.BGenereClick(Sender: TObject);
var
  Indice : integer;
begin
  if Regul=pglRegul then Indice:=3 else
  if Regul=pglEcart then Indice:=4
                    else Indice:=14;
  if TPieces.Count>0 then begin
    VisuPiecesGenere(TPieces,EcrGen,Indice);
    bDejaVu := True;
  end;
end;

Initialization
  registerclasses ( [ TOF_CPLETREGUL ] ) ;
end.
