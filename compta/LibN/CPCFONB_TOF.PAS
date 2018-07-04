{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 14/04/2003
Modifié le ... : 14/04/2003
Description .. : Source TOF de la FICHE : CPCFONB ()
Suite ........ : ATTENTION : Non testé en mode 2/3
Mots clefs ... : TOF;CPCFONB
*****************************************************************}
Unit CPCFONB_TOF;

Interface

Uses
    StdCtrls,
    Controls,
    Classes,
{$IFDEF EAGLCLIENT}
    MaineAGL,    // AGLLanceFiche
    eMul,        // TFMul
    UtileAGL,    // LanceEtat
{$ELSE}
    db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
    FE_Main,     // AGLLanceFiche
    Mul,         // TFMul
    EdtREtat,    // LanceEtat
    HDB,         // THDBGrid
{$ENDIF}
    Saisie,      // TrouveEtLanceSaisie
    MulSMPUtil,  // SmpToStr, StrToSmp
    SaisUtil,    // TSorteLettreToStr, StrToTSorteLettre
    utilPGI,     // _Blocage
    LettUtil,    // InitTablesLibresTiers
    HTB97,       // TToolBarButton97,
    Ent1,        // VH, ExoToDates
    ParamDat,    // ParamDate
    SaisComm,    // ModifRibOBM
    CFONB,       // ExportCFONB
    DocRegl,     // LanceDocRegl
    ed_tools,    // VideListe
    LookUp,      // LookupValueExist
    forms,
    sysutils,
    ComCtrls,
    HCtrls,
    HEnt1,
    HMsgBox,
    UTOF,
    UTOB,
    uRecupSQLModele,
    ParamSoc,		//GetParamSocSecur YMO
    HQRY
    ;

procedure ExportCFONBBatch(ONB,ENC : boolean; SorteLettre : TSorteLettre; smp : tSuiviMP = smpAucun);

Type
  TOF_CPCFONB = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure OnAfterFormShow          ;
    procedure InitMsgBox;

    procedure BOuvrirClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure BChercheClick(Sender: TObject);

    procedure Auxiliaire1Change(Sender: TObject);
    procedure Auxiliaire2Change(Sender: TObject);
    procedure E_EXERCICEChange(Sender: TObject);
    procedure E_DATECOMPTABLEKeyPress(Sender: TObject; var Key: Char);
    procedure bRibClick(Sender: TObject);
    procedure cRIBClick(Sender: TObject);
    procedure cExportClick(Sender: TObject);
    procedure CTIDTICClick(Sender: TObject);
    procedure AuxiElipsisClick(Sender : TObject);

  private
{$IFDEF EAGLCLIENT}
    FListe          : THGrid;
{$ELSE}
    FListe          : THDBGrid;
{$ENDIF}
    Q               : THQuery ;
    BanqueGene      : String ;
    FormshowEnCours : Boolean ;
    smp             : TsuiviMP;
    HM              : THMsgBox;
    FEcranCaption   : string;{JP 07/06/07 : FQ 12084}

    function  GetLeOBM : TOBM;
    function  ExporteSelection : boolean;
    procedure ChargeWhereTraite;
    procedure CompleteTL(TL : Tlist);
    procedure FlagExportLettre(TL : Tlist);
    function  CoherBanque(TL : TList) : boolean;
    function  GetLeTL(NbLig : integer) : TList;
    function  GetWhereSQL : string ;

    {JP 07/06/07 : FQ 12084 : Affichage du cumul de la sélection}
    function  CalculSoldeSelection : Double;
    procedure AfficheSoldeSelection;
    procedure SelectAllClick(Sender : TObject);
    procedure FListeFlipSelection(Sender : TObject);
  public
    ONB,ENC : Boolean ;
    SorteLettre : TSorteLettre;
  end ;

Implementation
uses
  {$IFDEF MODENT1}
  ULibExercice,
  {$ENDIF MODENT1}
  {$IFDEF eAGLCLIENT}
  MenuOLX
  {$ELSE}
  MenuOLG
  {$ENDIF eAGLCLIENT}
  , ExtCtrls, UTofMulParamGen; {13/04/07 YMO F5 sur Auxiliaire }


procedure ExportCFONBBatch(ONB,ENC : boolean; SorteLettre : TSorteLettre; smp : tSuiviMP = smpAucun);
var
  szArg : String;
begin
  if _Blocage(['nrCloture','nrBatch','nrLettrage'],True,'nrAucun') then Exit;
  if ONB then szArg := 'X;' else szArg := '-;';
  if ENC then szArg := szArg + 'X;' else szArg := szArg + '-;';
  szArg := szArg + TSorteLettreToStr(SorteLettre) + ';';
  szArg := szArg + SmpToStr(smp);
  AGLLanceFiche('CP','CPCFONB','','', szArg);
end;

procedure TOF_CPCFONB.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPCFONB.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPCFONB.OnUpdate ;
begin
  inherited;
  AfficheSoldeSelection; {JP 07/06/07 : FQ 12084}
end;

procedure TOF_CPCFONB.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPCFONB.OnArgument (S : String ) ;
begin
  Inherited ;
  // Création des contrôles
  HM := THMsgBox.Create(FMenuG);
  InitMsgBox;

  // Récupère les arguments
  SetControlText('XX_WHERE1', 'E_JOURNAL="zzz"');
  ONB := (ReadTokenSt(S) = 'X');
  ENC := (ReadTokenSt(S) = 'X');
  SorteLettre := StrToTSorteLettre(ReadTokenSt(S));
  Smp := StrToSmp(ReadTokenSt(S));
  Case SorteLettre of
    tslBOR    : TFMul(Ecran).SetDBListe('MULBOR');
    tslTraite : TFMul(Ecran).SetDBListe('MULTRAITE');
    else TFMul(Ecran).SetDBListe('MULCFONB');
  end;

  // Evénements des contrôles
  TToolBarButton97(GetControl('BOUVRIR',True)).OnClick   := BOuvrirClick;

  Q      := THQuery(GetControl('Q', True));
{$IFDEF EAGLCLIENT}
  FListe := THGrid(GetControl('FLISTE',True));
{$ELSE}
  FListe := THDBGrid(GetControl('FLISTE',True));
{$ENDIF}
  FListe.OnDblClick := FListeDblClick;
  TFMul(Ecran).OnAfterFormShow := OnAfterFormShow;

  TToolBarButton97(GetControl('BCHERCHE',True)).OnClick  := BChercheClick;

  THEdit(GetControl('AUXILIAIRE1',True)).OnChange        := Auxiliaire1Change;
  THEdit(GetControl('E_GENERAL',True)).OnChange          := Auxiliaire1Change;
  THEdit(GetControl('AUXILIAIRE2',True)).OnChange        := Auxiliaire2Change;
  THValComboBox(GetControl('E_EXERCICE',True)).OnChange  := E_EXERCICEChange;
  THEdit(GetControl('E_DATECOMPTABLE',True)).OnKeyPress  := E_DATECOMPTABLEKeyPress;
  THEdit(GetControl('E_DATECOMPTABLE_',True)).OnKeyPress := E_DATECOMPTABLEKeyPress;
  THEdit(GetControl('E_DATEECHEANCE',True)).OnKeyPress   := E_DATECOMPTABLEKeyPress;
  THEdit(GetControl('E_DATEECHEANCE_',True)).OnKeyPress  := E_DATECOMPTABLEKeyPress;
  TToolBarButton97(GetControl('BRIB',True)).OnClick      := bRibClick;
  TCheckBox(GetControl('CRIB',True)).OnClick             := cRIBClick;
  TCheckBox(GetControl('CEXPORT',True)).OnClick          := cExportClick;
  TCheckBox(GetControl('CTIDTIC',True)).OnClick          := CTIDTICClick;

  // FormShow
  FormShowEnCours := True;
  SetControlText('E_DEVISE', V_PGI.DevisePivot);
//  if ONB then SetControlEnabled('E_DEVISE', False); // 14741
  if (VH^.CPExoRef.Code<>'') then begin
    SetControlText('E_EXERCICE',       VH^.CPExoRef.Code);
    E_EXERCICEChange(Nil);
    SetControlText('E_DATECOMPTABLE',  DateToStr(VH^.CPExoRef.Deb));
    SetControlText('E_DATECOMPTABLE_', DateToStr(VH^.CPExoRef.Fin));
    end
  else begin
    SetControlText('E_EXERCICE',       VH^.Entree.Code);
    E_EXERCICEChange(Nil) ;
    SetControlText('E_DATECOMPTABLE',  DateToStr(V_PGI.DateEntree));
    SetControlText('E_DATECOMPTABLE_', DateToStr(V_PGI.DateEntree));
  end;
  SetControlText('E_DATEECHEANCE',     StDate1900);
  SetControlText('E_DATEECHEANCE_',    StDate2099);

  InitTablesLibresTiers(TTabSheet(GetControl('PLIBRES',True)));
  SetControlText('XX_WHERE1', '');
  BanqueGene := '';

  if (SorteLettre <> tslAucun) then begin
    SetControlVisible('H_Document', True);
    SetControlVisible('DOCUMENT',   True);
    SetControlVisible('APERC',      True);

    SetControlProperty('DOCUMENT','DATATYPE', 'TTMODELELETTRETRA');
    ChargeComboEtat(THValCOmboBox(GetControl('DOCUMENT')),'L','LTR','') ;  //XVI FQ 15425 (non corrigé encore)
    SetControlText('H_DOCUMENT', HM.Mess[9]); // Modèle

    SetControlProperty('E_MODEPAIE','DATATYPE', 'TTMODEPAIETRAITE');

    ChargeWhereTraite;

    SetControlVisible('APERC',   False);
    SetControlChecked('CEXPORT', True);
    SetControlText('XX_WHERE1', 'E_ETATLETTRAGE<>"RI" AND E_LETTRAGE<>""');
    if (SorteLettre = tslBOR) then begin
      SetControlText('XX_WHERENAT', 'E_NATUREPIECE="OD" OR E_NATUREPIECE="RF" OR E_NATUREPIECE="AF" OR E_NATUREPIECE="OF"');
      SetControlText('XX_WHEREDC',  'E_DEBIT>0 OR E_CREDIT<0');
      FEcranCaption := HM.Mess[7]; // Emission de lettres-BOR
      end
    else begin
      SetControlText('XX_WHERENAT', 'E_NATUREPIECE="OD" OR E_NATUREPIECE="RC" OR E_NATUREPIECE="AC" OR E_NATUREPIECE="OC"');
      SetControlText('XX_WHEREDC',  'E_DEBIT<0 OR E_CREDIT>0');
      FEcranCaption := HM.Mess[11]; // Emission de lettres-traite
    end;
    end
  else if Not ONB then begin
    SetControlVisible('H_DOCUMENT', True);
    ChargeComboEtat(THValCOmboBox(GetControl('DOCUMENT')),'E','BOR','') ;  //XVI FQ 15425 (non corrigé encore)
    SetControlVisible('DOCUMENT', True);
    SetControlVisible('APERC', True);

    SetControlProperty('E_JOURNAL', 'DATATYPE', '');
    SetControlProperty('E_JOURNAL', 'VIDE', False);
    SetControlProperty('E_JOURNAL', 'DATATYPE', 'TTJALSANSECART');
    if (THValComboBox(GetControl('E_JOURNAL',True)).Items.Count > 0) then
      THValComboBox(GetControl('E_JOURNAL')).ItemIndex := 0;

    {JP 08/06/07 : FQ 19255 : on se positionne sur le premier}
    if (THValComboBox(GetControl('DOCUMENT')).Items.Count > 0) then
      THValComboBox(GetControl('DOCUMENT')).ItemIndex := 0;

    FEcranCaption := HM.Mess[4]; // Emission de bordereaux
    end
  else begin
    if ENC then begin FEcranCaption := HM.Mess[5]; SetControlText('XX_WHEREDC', ''); end  // Export CFONB des encaissements
           else begin FEcranCaption := HM.Mess[6]; SetControlText('XX_WHEREDC', ''); end; // Export CFONB des décaissements
  end;

  If ONB And ENC then Ecran.HelpContext := 7586000 Else
    If (Not ONB) And ENC then Ecran.HelpContext := 999999824 Else
      If (ONB) And (Not ENC) then Ecran.HelpContext := 7595000;

  if (SorteLettre = tslTraite) then Ecran.HelpContext := 7595200 else
    if (SorteLettre = tslBOR) then Ecran.HelpContext := 7595400 ;

//  if ONB And (SorteLettre=tslAucun) And (Not Document.Visible) then SetControlVisible('CTIDTIC', True);

  if (SorteLettre=tslAucun) then SetControlVisible('CTIDTIC', True);
{$IFDEF CCMP}
  if (SorteLettre = tslAucun) And (Not ONB) then CTIDTICClick(Nil);
  if (SorteLettre = tslAucun) Then begin
    if ONB And Enc then Ecran.HelpContext := 7586000;
    if ONB And (Not Enc) then Ecran.HelpContext := 7595000;
    if (Not ONB) Then Ecran.HelpContext := 7598210;
  end;
{$ENDIF}

  Ecran.Caption := FEcranCaption; {JP 07/06/07 : FQ 12084}
  UpdateCaption(Ecran);

  TToolBarButton97(GetControl('BSELECTALL', True)).OnClick := SelectAllClick;{JP 07/06/07 : FQ 12084}
  FListe.OnFlipSelection := FListeFlipSelection;{JP 07/06/07 : FQ 12084}

  if GetParamSocSecur('SO_CPMULTIERS', false) then
  begin
    THEdit(GetControl('AUXILIAIRE1', true)).OnElipsisClick:=AuxiElipsisClick;
    THEdit(GetControl('AUXILIAIRE2', true)).OnElipsisClick:=AuxiElipsisClick;
  end;

  {JP 30/10/07 : FQ 21432 : Glyph du bouton valider "VVVert"}
  TFMul(Ecran).BOuvrir.GlobalIndexImage := 'Z0184_S16G1';
end ;

procedure TOF_CPCFONB.OnClose ;
begin
  Inherited ;
  HM.Free;
end ;

procedure TOF_CPCFONB.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPCFONB.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPCFONB.InitMsgBox;
begin
  HM.Mess.Add('0;?Caption?;Export ou émission impossible. Vous n''avez rien sélectionné.;E;O;O;O;');
  HM.Mess.Add('1;?Caption?;Confirmez-vous l''export ou l''émission des échéances sélectionnées ?;Q;YN;Y;Y;');
  HM.Mess.Add('2;?Caption?;ATTENTION ! Vous avez tout sélectionné. Confirmez-vous l''export ou l''émission ?;W;YN;Y;Y;');
  HM.Mess.Add('3;?Caption?;Vous devez sélectionner un modèle de document pour l''émission du bordereau.;W;O;O;O;');
  HM.Mess.Add('Emission de bordereaux');
  HM.Mess.Add('Export CFONB des encaissements');
  HM.Mess.Add('Export CFONB des décaissements');
  HM.Mess.Add('Emission de lettres-BOR');
  HM.Mess.Add('8;?Caption?;Vous devez sélectionner un modèle de document.;W;O;O;O;');
  HM.Mess.Add('Modèle');
  HM.Mess.Add('10;?Caption?;Vous devez renseigner un compte général.;W;O;O;O;');
  HM.Mess.Add('Emission de lettres-traite');
  HM.Mess.Add('12;?Caption?;ATTENTION ! Vous avez des banques différentes. L''export sera incohérent. Confirmez-vous l''export ou l''émission ?;W;YN;Y;Y;');
end;

procedure TOF_CPCFONB.Auxiliaire1Change(Sender: TObject);
begin
  SetControlText('E_AUXILIAIRE', GetControlText('AUXILIAIRE1'));
end;

procedure TOF_CPCFONB.Auxiliaire2Change(Sender: TObject);
begin
  SetControlText('E_AUXILIAIRE_', GetControlText('AUXILIAIRE2'));
end;

procedure TOF_CPCFONB.bRibClick(Sender: TObject);
var O     : TOBM ;
    IsAux : Boolean ;
begin

  if GetDataSet.EOF then Exit ;

  O := GetLeOBM;
  if O<>Nil then
    begin
    IsAux := O.GetString('E_AUXILIAIRE')<>'';
    if ModifRibOBM(O, True, False, '', IsAux) then
      bChercheClick(nil) ;
    FreeANdNil( O );
    end ;

end;

procedure TOF_CPCFONB.cExportClick(Sender: TObject);
begin
  case TCheckBox(GetControl('CEXPORT', True)).State of
    cbGrayed    : SetControlText('XX_WHEREEXPORT', '');
    cbChecked   : SetControlText('XX_WHEREEXPORT', 'E_CFONBOK="X"');
    cbUnchecked : SetControlText('XX_WHEREEXPORT', 'E_CFONBOK<>"X"');
  end;
end;

procedure TOF_CPCFONB.cRIBClick(Sender: TObject);
begin
  case TCheckBox(GetControl('CRIB', True)).State of
    cbGrayed    : SetControlText('XX_WHERERIB', '');
    cbChecked   : SetControlText('XX_WHERERIB', 'E_RIB<>""');
    cbUnchecked : SetControlText('XX_WHERERIB', 'E_RIB="" or E_RIB="////"');
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 12/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCFONB.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;


procedure TOF_CPCFONB.CTIDTICClick(Sender: TObject);
var
  b : Boolean;
begin
  b := not TCheckBox(GetControl('CTIDTIC')).Checked;
  SetControlVisible('AUXILIAIRE1',  b);
  SetControlVisible('AUXILIAIRE2',  b);
  SetControlVisible('HAUXILIAIRE1', b);
  SetControlVisible('HAUXILIAIRE2', b);
  if not b then begin
    SetControlText('AUXILIAIRE1', '');
    SetControlText('AUXILIAIRE2', '');
    If ENC then SetControlProperty('E_GENERAL', 'DATATYPE', 'tzGTID')
           else SetControlProperty('E_GENERAL', 'DATATYPE', 'tzGTIC');
    SetControlText('E_GENERAL', '');
    SetControlText('XX_WHEREAUX', 'E_AUXILIAIRE="" AND E_NUMECHE>0 AND E_ETATLETTRAGE<>"RI" ');
    end
  else begin
{$IFDEF CCMP}
    if Enc then begin
      SetControlProperty('E_GENERAL',   'DATATYPE', 'tzGCollClient');
      SetControlProperty('AUXILIAIRE1', 'DATATYPE', 'tzTToutDebit');
      SetControlProperty('AUXILIAIRE2', 'DATATYPE', 'tzTToutDebit');
      end
    else begin
      SetControlProperty('E_GENERAL',   'DATATYPE', 'tzGCollFourn');
      SetControlProperty('AUXILIAIRE1', 'DATATYPE', 'tzTToutCredit');
      SetControlProperty('AUXILIAIRE2', 'DATATYPE', 'tzTToutCredit');
    end;
{$ELSE}
    SetControlProperty('E_GENERAL', 'DATATYPE', 'tzGCollectif');
{$ENDIF}
    SetControlText('XX_WHEREAUX', 'E_AUXILIAIRE<>""');
  end;
end;

procedure TOF_CPCFONB.E_DATECOMPTABLEKeyPress(Sender: TObject; var Key: Char);
begin
  ParamDate(Ecran,Sender,Key);
end;

procedure TOF_CPCFONB.E_EXERCICEChange(Sender: TObject);
begin
  ExoToDates(GetControlText('E_EXERCICE'), GetControl('E_DATECOMPTABLE',True), GetControl('E_DATECOMPTABLE_',True) );
end;

procedure TOF_CPCFONB.ChargeWhereTraite;
var
  QQ : TQuery ;
  St : String ;
begin
  QQ := OpenSQL('SELECT MP_MODEPAIE FROM MODEPAIE WHERE MP_LETTRETRAITE="X"',True);
  St := '';
  While Not QQ.EOF do begin
    St := St + 'E_MODEPAIE="' + QQ.Fields[0].AsString + '" OR ';
    QQ.Next;
  end;
  Ferme(QQ);
  if (St <> '') then System.Delete(St,Length(St)-2,3);
  SetControlText('XX_WHERE1', St);
end;

function TOF_CPCFONB.CoherBanque(TL: TList): boolean;
var
  i : integer ;
  O : TOBM ;
  Jal,OldJal : String3 ;
  Okok : boolean ;
  QQ   : TQuery ;
begin
  if ENC then begin Result := True; Exit; end;
  Okok := True;
  OldJal := '';
  for i := 0 to TL.Count-1 do begin
    O := TOBM(TL[i]);
    if O=Nil then Break;
    Jal := O.GetString('E_JOURNAL');
    if ((OldJal<>'') and (Jal<>OldJal)) then begin Okok := False; Break; end;
    OldJal := Jal;
  end;

  if Not Okok then begin
    // Vous avez des banques différentes. L'export sera incohérent. Confirmez-vous l'export ou l'émission ?
    Okok := (HM.Execute(12, FEcranCaption,'')=mrYes);
    end
  else if OldJal<>'' then begin
    QQ := OpenSQL('SELECT J_CONTREPARTIE FROM JOURNAL WHERE J_JOURNAL="'+OldJal+'"',True);
    if Not QQ.EOF then BanqueGene := QQ.Fields[0].AsString;
    Ferme(QQ);
  end;
  Result := Okok;
end;

procedure TOF_CPCFONB.CompleteTL(TL: Tlist);
var
  i : integer ;
  O,O2 : TOBM ;
  RR   : RMVT ;
  QL : TQuery ;
  CodeL,SQL,St : String ;
begin
  {JP 09/05/07 : C'est plus prudent !!!}
  if not Assigned(TL) then Exit;

  for i := 0 to TL.Count-1 do begin
    O:=TOBM(TL[i]);
    if O=Nil then Break ;
    CodeL := O.GetString('E_LETTRAGE');
    if CodeL='' then Continue;
    if O.GetString('E_ETATLETTRAGE')='RI' then Continue;
    SQL := 'SELECT * FROM ECRITURE WHERE E_AUXILIAIRE="' + O.GetString('E_AUXILIAIRE') + '"'
         + 'AND E_GENERAL="' + O.GetString('E_GENERAL') + '" AND E_LETTRAGE="' + CodeL + '"';
    QL := OpenSQL(SQL,True);
    While Not QL.EOF do begin
      O2:=TOBM.Create(EcrGen,'',False);
      O2.ChargeMvt(QL) ;

      // Ne pas se reprendre elle-même
      if ((O2.GetInteger('E_NUMEROPIECE')    <> O.GetInteger('E_NUMEROPIECE')) or
          (O2.GetInteger('E_NUMLIGNE')       <> O.GetInteger('E_NUMLIGNE'))    or
          (O2.GetInteger('E_NUMECHE')        <> O.GetInteger('E_NUMECHE'))     or
          (O2.GetString('E_JOURNAL')         <> O.GetString('E_JOURNAL'))      or
          (O2.GetDateTime('E_DATECOMPTABLE') <> O.GetDateTime('E_DATECOMPTABLE'))) then begin
        RR := OBMToIdent(O2,True);
        St := EncodeLC(RR);
        O2.Free;
        O.LC.Add(St);
        {T:=TStringList.Create;
        T.Add(St);
        O.LC.Assign(T);
        T.Free;}
        end
      else begin
        O2.Free ;
      end;
      QL.Next;
    end;
    Ferme(QL);
  end;
end;

(*
function TOF_CPCFONB.ExporteSelection: boolean;
var
  i,NbLig,ii : integer ;
  TL : TList ;
  O       : TOBM ;
  SWhere : String ;
  Inutile : TMSEncaDeca ;
  Q : TQuery;
{$IFDEF EAGLCLIENT}
  Fiche : TFMul;
{$ENDIF}
begin
{$IFDEF EAGLCLIENT}
  Q := TFMul(Ecran).Q.TQ;
  Q.Seek(TFMul(Ecran).FListe.Row-1);
{$ELSE}
  Q := TFMul(Ecran).Q;
{$ENDIF}
  Result := False;
  Fillchar(Inutile,SizeOf(Inutile),#0);
  if ((Not ONB) and (GetControlText('DOCUMENT') = '')) then begin
    if (SorteLettre=tslAucun) then HM.Execute(3,Ecran.Caption,'')  // Vous devez sélectionner un modèle de document pour l'émission du bordereau
                              else HM.Execute(8,Ecran.Caption,''); // Vous devez sélectionner un modèle de document.
    Exit ;
  end;

  if Not TFMul(Ecran).FListe.AllSelected then begin
    NbLig := TFMul(Ecran).FListe.NbSelected ;
    if (NbLig <= 0) then begin
      HM.Execute(0,Ecran.Caption,''); // Export ou émission impossible. Vous n'avez rien sélectionné.
      Exit;
    end;

    // Confirmez-vous l'export ou l'émission des échéances sélectionnées ?
    if HM.Execute(1,Ecran.Caption,'')<>mrYes then Exit;

    {$IFDEF EAGLCLIENT}
    Fiche := TFMul(Ecran);
    if Fiche.bSelectAll.Down then
      if not Fiche.FetchLesTous then
      begin
        Fiche.bSelectAllClick(nil);
        Fiche.bSelectAll.Down := False;
        Exit;
      end;
    {$ENDIF}

    TL := Tlist.Create;
    for i:=0 to NbLig-1 do begin
      TFMul(Ecran).FListe.GotoLeBookMark(i);
      O := GetLeOBM;
      if O<>Nil then TL.Add(O);
    end;
    end
  else begin
    // Vous avez tout sélectionné. Confirmez-vous l'export ou l'émission ?
    if HM.Execute(2,Ecran.Caption,'')<>mrYes then Exit;
    
    TL := Tlist.Create;
    Q.First;
    ii:=0;
    While ((Not Q.EOF) and (ii<1000)) do begin
      O := GetLeOBM;
      if O<>Nil then TL.Add(O);
      Q.Next;
      Inc(ii);
    end;
  end;

  if ONB then begin
    if CoherBanque(TL) then ExportCFONB(ENC,BanqueGene,'','DEM',TL);
    end
  else if SorteLettre<>tslAucun then begin
    CompleteTL(TL);
    if LanceDocRegl(TL, SorteLettre, GetControlText('DOCUMENT'), '', False, Inutile)>0 then FlagExportLettre(TL);
    end
  else begin
    SWhere := WhereMulti(TL) ;
    LanceEtat('E', 'BOR', GetControlText('DOCUMENT'), TCheckBox(GetControl('APERC')).Checked, TCheckBox(GetControl('FEXPORT')).Checked, False, Nil, SWhere, '', False);
  end;
  VideListe(TL);
  TL.Free;
  Result := True;
  BChercheClick(Nil);
end;
*)

// BPY le 17/02/2004 => fonction pour obtenir le TL ....
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 17/02/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function  TOF_CPCFONB.GetLeTL(NbLig : integer) : TList;
var
    TL : TList;
    i : integer;
    O : TOBM;
begin

    result := nil;

    {$IFNDEF EAGLCLIENT}
    {JP 23/10/07 : FQ En CWAS on fait toujours la boucle FOR car GetDataSet.Next n'est pas compatible eAGL}
    // recup de chaque ligne !
    if (not FListe.AllSelected) then
    begin
    {$ENDIF EAGLCLIENT}
        if (HM.Execute(1, FEcranCaption,'') <> mrYes) then exit;

        TL := Tlist.Create;
        for i := 0 to NbLig-1 do
        begin
            FListe.GotoLeBookMark(i);
            O := GetLeOBM;
            if (O <> Nil) then TL.Add(O);
        end;
    {$IFNDEF EAGLCLIENT}
    end
    else
    begin
        if (HM.Execute(2, FEcranCaption,'') <> mrYes) then exit;

        TL := Tlist.Create;

        GetDataSet.First ;
        while (Not GetDataSet.EOF) do
          begin
          O := GetLeOBM;
          if (O <> Nil) then TL.Add(O);
          GetDataSet.Next;
          end;
    end;
    {$ENDIF EAGLCLIENT}

    result := TL;
end;

// BPY le 08/10/2003
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 08/10/2003
Modifié le ... : 20/02/2004
Description .. : 2003-02-20 : Suppression de la limite a 1000 Lignes ....
Suite ........ :     en passant par une TOB et un LanceEtatTob
Mots clefs ... : 
*****************************************************************}
function TOF_CPCFONB.ExporteSelection : boolean ;
var
    i,NbLig : integer;
    TL : TList;
    LO : TStringList;
    TE : TOB;
    SQuery,SWhere,AWhere,SGroup,SOrder : String;
    Inutile : TMSEncaDeca;
begin
    result := false;
    Fillchar(Inutile,SizeOf(Inutile),#0);

    // pas de document selectionné .... out !
    if ((not ONB) and (GetControlText('DOCUMENT') = '')) then
    begin
        if (SorteLettre = tslAucun) then HM.Execute(3, FEcranCaption,'')
        else HM.Execute(8, FEcranCaption,'');
        exit;
    end;

    // Recup du nb de Ligne
{$IFDEF EAGLCLIENT}
    NbLig := FListe.NbSelected;
{$ELSE}
    if (FListe.AllSelected)
      then NbLig := GetDataSet.RecordCount
      else NbLig := FListe.NbSelected ;
{$ENDIF}

    // si pas de lignes ...
    if (NbLig <= 0) then
    begin
        HM.Execute(0, FEcrancaption,'');
        exit;
    end;

    // traitement
    if (ONB) then
    begin
        TL := GetLeTL(NbLig);
        {JP 09/05/07 : FQ 19609 :C'est plus prudent, car GetLeTL demande si l'on veut poursuivre?
                       Si la réponse est non ...}
        if not Assigned(TL) then Exit;

        if (CoherBanque(TL)) then ExportCFONB(ENC,BanqueGene,'','DEM',TL);
        VideListe(TL);
        TL.Free;
    end
    else if (SorteLettre <> tslAucun) then
    begin
        TL := GetLeTL(NbLig);
        {JP 09/05/07 : FQ 19609 :C'est plus prudent, car GetLeTL demande si l'on veut poursuivre?
                       Si la réponse est non ...}
        if not Assigned(TL) then Exit;

        CompleteTL(TL);
        if (LanceDocRegl(TL,SorteLettre,GetControlText('DOCUMENT'),'',False,Inutile) > 0) then FlagExportLettre(TL);
        VideListe(TL);
        TL.Free;
    end
    else
    begin
    // BPY le 16/02/2004 => depassement de la limite a 1000 ligne !
//        TL := GetLeTL(NbLig);
//        SWhere := WhereMulti(TL);
//        LanceEtat('E','BOR',Document.Value,Aperc.Checked,FExport.Checked,False,Nil,SWhere,'',False);
//        VideListe(TL);
//        TL.Free;
        TE := TOB.Create('',nil,-1);
        SQuery := RecupSQLComplet('E','BOR',GetControlText('DOCUMENT'),SWhere,SGroup,SOrder);

        // recuperation d'une tob pour chaque line !
        // !!! ATTENTION !!! => La clause goup by ne sera pas prise en compte !
{$IFDEF EAGLCLIENT}
        if (not (NbLig = GetDataSet.RecordCount)) then
{$ELSE}
        if (not (FListe.AllSelected or (NbLig = GetDataSet.RecordCount))) then
{$ENDIF}
        begin
            if (HM.Execute(1, FEcranCaption,'') <> mrYes) then exit;
            for i := 0 to NbLig-1 do
            begin
                FListe.GotoLeBookMark(i);
                TE.LoadDetailFromSQL( SQuery + ' ' + SWhere + GetWhereSQL, true ) ;
            end;

            // trie de la tob !
            if (not (SOrder = '')) then
            begin
                SOrder := copy(SOrder,10,length(SOrder)-10); // Supprime le "group by "
                LO := TStringList.Create();
                while (not (SOrder = '')) do LO.Add(ReadTokenPipe(SOrder,','));
                for i := LO.Count-1 downto 0 do TE.Detail.Sort(LO[i]);
                FreeAndNil(LO);
            end;
        end
        else
        begin
            if (HM.Execute(2, FEcranCaption,'') <> mrYes) then exit;
            AWhere := RecupWhereCritere(TFMul(Ecran).Pages);
            AWhere := 'AND ' + copy(AWhere,6,length(Awhere)-6);
            TE.LoadDetailFromSQL(SQuery + ' ' + SWhere + ' ' + AWhere + ' ' + SOrder);
        end;

        // lancement de l'etat
        // !!! ATTENTION !!! => ne marche pas s'il y a des bande report dans l'etat !
        LanceEtatTob('E','BOR',GetControlText('DOCUMENT'),TE,TCheckBox(GetControl('APERC')).Checked,TCheckBox(GetControl('FEXPORT')).Checked,false,nil,'','',false);
        FreeAndNil(TE);
    // fin BPY
    end;

    result := true;
    BChercheClick(nil);
end;
// Fin BPY


procedure TOF_CPCFONB.FlagExportLettre(TL: Tlist);
var
  i : integer ;
  O : TOBM ;
  MM : RMVT ;
begin
  for i := 0 to TL.Count-1 do begin
    O := TOBM(TL[i]);
    if (O = Nil) then Exit;
    MM := OBMToIdent(O,True);
    ExecuteSQL('UPDATE ECRITURE SET E_CFONBOK="X" WHERE '+WhereEcriture(tsGene,MM,True));
  end;
end;

function TOF_CPCFONB.GetLeOBM: TOBM;
var Q1 : TQuery ;
    O  : TOBM ;
begin
  O:=Nil;

  Q1:=OpenSQL('SELECT * FROM ECRITURE WHERE ' + GetWhereSQL , True ) ;

  if (Not Q1.EOF) then begin
    O := TOBM.Create(EcrGen,'',False);
    O.ChargeMvt(Q1);
  end;
  Ferme(Q1);
  Result := O;
end;

procedure TOF_CPCFONB.BChercheClick(Sender: TObject);
begin
  {JP 09/05/07 : FQ 19609 : En V8 FormShowEnCours semble fonctionner, alors que ce n'était pas le cas en V7.
                 L'agl a dû bouger. Pour éviter toute surprise avec OG, on supprime le test sur la recherce
  if (Not FormShowEnCours) and ((SorteLettre <> tslAucun) and (not LookupValueExist(GetControl('E_GENERAL')))) then
    begin
    HM.Execute(10,Ecran.Caption,''); // Vous devez renseigner un compte général.
    if THEdit(GetControl('E_GENERAL',True)).CanFocus then
      SetFocusControl('E_GENERAL');
    SetControlEnabled('BRIB', False);
    Exit ;
    end;
  }
  if (SorteLettre <> tslAucun) and not LookupValueExist(GetControl('E_GENERAL')) then
    SetControlEnabled('BRIB', False);

  TFMul(Ecran).BChercheClick(Sender); // inherited;

  SetControlEnabled('BRIB', not (GetDataSet.EOF));
end;

procedure TOF_CPCFONB.BOuvrirClick(Sender: TObject);
begin
  if ( (SorteLettre <> tslAucun) and (not LookupValueExist(GetControl('E_GENERAL'))) ) then begin
    HM.Execute(10, FEcranCaption,''); // Vous devez renseigner un compte général.
    if THEdit(GetControl('E_GENERAL',True)).CanFocus then SetFocusControl('E_GENERAL');
    Exit ;
  end;
  if Not ExporteSelection then Exit ;

  if Not FListe.AllSelected
    then FListe.ClearSelected
    else FListe.AllSelected := False;
end;

procedure TOF_CPCFONB.FListeDblClick(Sender: TObject);
begin
  if GetDataSet.EOF then Exit ;
{$IFDEF EAGLCLIENT}
  Q.TQ.Seek(FListe.Row-1);
  TrouveEtLanceSaisie(Q.TQ,taConsult,GetControlText('E_QUALIFPIECE'));
{$ELSE}
  TrouveEtLanceSaisie(Q,taConsult,GetControlText('E_QUALIFPIECE'));
{$ENDIF}
end;

function TOF_CPCFONB.GetWhereSQL: string;
begin
result := ' AND E_JOURNAL="'        + GetField('E_JOURNAL')
       + '" AND E_EXERCICE="'       + QuelExo(DateToStr(GetField('E_DATECOMPTABLE')))
       + '" AND E_DATECOMPTABLE="'  + USDATETIME(GetField('E_DATECOMPTABLE'))
       + '" AND E_NUMEROPIECE='     + IntToStr( GetField('E_NUMEROPIECE') )
       + ' AND E_NUMLIGNE='         + IntToStr( GetField('E_NUMLIGNE') )
       + ' AND E_QUALIFPIECE="N"'
       + ' AND E_NUMECHE='          + IntToStr( GetField('E_NUMECHE') );
end;

procedure TOF_CPCFONB.OnAfterFormShow;
begin
  FormShowEnCours := False;
end;

{JP 07/06/07 : FQ 12084 : Affichage du cumul de la sélection
{---------------------------------------------------------------------------------------}
procedure TOF_CPCFONB.AfficheSoldeSelection;
{---------------------------------------------------------------------------------------}
var
  Solde : Double;
begin
  {Calcul du solde}
 	Solde := CalculSoldeSelection;
  {Ajout affichage du cumul}
  if Solde < 0 then
  	 Ecran.Caption := FEcranCaption + ' (Cumul de la sélection : ' + StrFMontant(-1 * Solde, 0, V_PGI.OkDecV, '', True) + ' C)'
  else
    Ecran.Caption := FEcranCaption + ' (Cumul de la sélection : ' + StrFMontant(Solde, 0, V_PGI.OkDecV, '', True) + ' D)';

	{Raffraichissement Titre}
	UpdateCaption(Ecran) ;
end;

{JP 07/06/07 : FQ 12084 : Affichage du cumul de la sélection
{---------------------------------------------------------------------------------------}
function TOF_CPCFONB.CalculSoldeSelection : Double;
{---------------------------------------------------------------------------------------}

    {-------------------------------------------------------------}
    function _GetCumul : Double;
    {-------------------------------------------------------------}
    var
      n : Integer;
    begin
      Result := 0;

      for n := 0 to FListe.NbSelected - 1 do begin
        FListe.GotoLeBookmark(n);
        Result := Result + Valeur(GetField('E_DEBIT')) - Valeur(GetField('E_CREDIT'));
      end;
    end;

var
	Debit  : Double;
  Credit : Double;
  PCumul : TPanel;
begin
  if FListe.AllSelected then
  begin
    PCumul := TPanel(GetControl('PCumul', True));

    if (PCumul.FindChildControl('__QRYPCumul_E_DEBIT') <> nil) and
       (PCumul.FindChildControl('__QRYPCumul_E_CREDIT') <> nil) then begin

      Debit  := Valeur(THNumEdit(PCumul.FindChildControl('__QRYPCumul_E_DEBIT')).Text);
      Credit := Valeur(THNumEdit(PCumul.FindChildControl('__QRYPCumul_E_CREDIT')).Text);
      Result := Debit - Credit;
    end
    else
      Result := _GetCumul;
  end
  else
    Result := _GetCumul;
end;

{JP 07/06/07 : FQ 12084 : Affichage du cumul de la sélection
{---------------------------------------------------------------------------------------}
procedure TOF_CPCFONB.SelectAllClick(Sender : TObject);
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

 	AfficheSoldeSelection;
end;

{JP 07/06/07 : FQ 12084 : Affichage du cumul de la sélection
{---------------------------------------------------------------------------------------}
procedure TOF_CPCFONB.FListeFlipSelection(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
 	AfficheSoldeSelection;
end;

Initialization
  registerclasses ( [ TOF_CPCFONB ] ) ;
end.
